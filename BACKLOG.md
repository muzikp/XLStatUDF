# Backlog - PRIORITNÍ

## Nefunknční UDF

- aktuálně v pořádku

## UI

### různé

- NORM.DIST.RANGE dát do sekce "popisné" (již existuje)
- UDF v dokumentaci seřadit (uvnitř sekcí) abecedne
- u CONTINGENCY.G funkce renderovat ve výstupu phi jako řecké písmeno

### dokumentace

- stále nacházím občasné rozdíly mezi dokumentací v sidepanelu a v samotném výstupu UDF (možná i parametrech UDF), mělo by to být totožné
- dbej i na formální úpravu, tj. stejné formátování popisků, včetně využívání písmen řecké abecedy, dolních a horních indexů a dalších speciálních znaků
- buď jednotný v popisech, např. hodnoty kritických oborů můžeme popisovat jako název testu a dolním indexem počítanou pravděpodobnost, mělo by to být ale napříč řešením stejné, pokud k tomu nemáš zvláštní důvod to dělat různě

# Backlog - K DISKUSI

## Autorizace

- potřebuji kontrolovat, kdo bude knihovnu využívat, eventuálně kdy a jaké funkce bude mít k dispozici
- nástroj je aktuálně určen zejména studentům ČVUT, v budoucnu možná VŠE, eventuálně bude mít přístup i veřejnost (přes další autorizační služby typu email/password, OAuth Google, GitHub); při uvažování architektury bude dobré zvážit vše, co tu píšu, ale primárně mi půjde o vývoj autorizace přes ČVUT OAuth
- pokud nebude uživatel přihlášen, funkce mu budou vracet nějakou chybou hodnotu (např. #UNAUTHORIZED)
- bylo by asi vhodné i logovat vybrané události; k tomu můžeme použít AWS (např. Lambda + MySQL RDS, obojí poskytnu), kde by se zpracovaly další práva (blacklist, omezení funkcí apod.)


### Navrh implementace autorizace

#### Cile

- Primarni cil je ridit pristup k Excel add-inu a jednotlivym UDF podle identity uzivatele.
- Prvni podporovany provider ma byt CVUT OAuth/OIDC. Architektura ma zustat pripravena na VSE, Google, GitHub a email/password.
- Neprihlaseny nebo zablokovany uzivatel nema dostat vysledek chranene UDF. Funkce ma vratit jasnou chybu, idealne #N/A/notAvailable s textem #UNAUTHORIZED nebo Unauthorized.
- Opravneni se maji vyhodnocovat centralne na serveru, ne jen v klientskem JS. Klientsky add-in je verejne citelny a nesmi obsahovat zadna tajemstvi.
- System ma logovat vybrane udalosti: prihlaseni, odhlaseni, refresh tokenu, pouziti UDF, odmitnuti UDF, chyby autorizace, zmeny opravneni.

#### Dulezite omezeni Office add-inu

- Taskpane umi pohodlne otevrit login flow pres Office dialog nebo nove okno.
- Custom functions runtime bezi oddelene od taskpane, ale v tomto projektu uz umi cist localStorage kvuli jazyku vystupu. Autorizaci proto lze sdilet pres stejny origin pomoci localStorage.
- Spolehat jen na localStorage neni bezpecnostni ochrana. Je to jen cache stavu pro UX. Skutecne rozhodnuti musi delat backend nebo backendem podepsany kratkodoby token.
- Pokud kazda UDF bude pri kazdem prepoctu volat backend, muze to byt pomale a krehke. Doporuceny kompromis je kratkodoby podepsany entitlement token s TTL 5-15 minut, plus serverove logovani davkovane nebo throttlovane.
- Presny zpusob vraceni chyby v UDF: pridat helper unauthorized() v office-addin/src/custom-functions/functions.ts, ktery vraci new CustomFunctions.Error(CustomFunctions.ErrorCode.notAvailable, "#UNAUTHORIZED"). Tim vznikne Excel chyba s vysvetlujicim textem, misto obycejneho stringu.

#### Doporucena architektura

- Frontend/add-in:
  - office-addin/src/public/taskpane.js: UI pro prihlaseni, stav uctu, odhlaseni, refresh.
  - office-addin/src/custom-functions/functions.ts: autorizacni helper pouzivany vsemi chranenymi UDF.
  - office-addin/src/public/auth-callback.html nebo podobna callback stranka pro dokonceni OAuth flow v dialogu.
- Backend v AWS:
  - API Gateway + Lambda pro auth API.
  - MySQL RDS pro uzivatele, providery, opravneni, blacklist, audit logy.
  - Volitelne AWS Secrets Manager pro OAuth client secret a JWT signing key.
  - Volitelne CloudWatch pro provozni logy a alarmy.
- Domeny:
  - Add-in a auth API by idealne mely bezet pod stejnou domenovou zonou, napriklad https://evalytics... pro add-in a https://api.evalytics... pro API.
  - Pro custom functions je praktictejsi pouzivat Bearer token v Authorization headeru nez spolehat na cookies, protoze runtime a WebView se mohou chovat odlisne.

#### Auth flow pro CVUT OAuth/OIDC

- Pred implementaci overit aktualni CVUT OAuth/OIDC dokumentaci a endpointy. Do kodu nedavat domnele endpointy natvrdo bez overeni.
- Doporuceny flow:
  1. Taskpane zavola GET /auth/providers a zjisti dostupne providery.
  2. Uzivatel klikne na Prihlasit pres CVUT.
  3. Taskpane otevre Office dialog na GET /auth/login/cvut/start.
  4. Backend vytvori state, nonce a PKCE verifier/challenge, ulozi kratkodoby login transaction.
  5. Backend presmeruje na CVUT authorization endpoint.
  6. CVUT vrati callback na GET /auth/login/cvut/callback.
  7. Backend vymeni code za tokeny, overi ID token / userinfo, normalizuje identitu.
  8. Backend zalozi nebo aktualizuje uzivatele, vyhodnoti opravneni, vytvori Evalytics session.
  9. Callback stranka posle vysledek zpet taskpane pres Office dialog message.
  10. Taskpane ulozi do localStorage kratkodoby evalytics.auth.accessToken, evalytics.auth.expiresAt, evalytics.auth.user, evalytics.auth.entitlements.
  11. Taskpane vyzada full recalculation, aby se UDF po prihlaseni prepocetly.
- Pro refresh:
  - Taskpane vola POST /auth/refresh pred expiraci.
  - Custom functions runtime se nesnazi otevirat login dialog. Kdyz token chybi nebo expiroval, vrati #UNAUTHORIZED.
  - Pokud UDF narazi na expirovany token, muze zkusit jeden tichy refresh jen v pripade, ze refresh mechanismus funguje bez UI. Jinak vracet #UNAUTHORIZED.

#### Entitlement token

- Backend vydava kratkodoby JWT nebo PASETO token pro add-in.
- Token nema obsahovat citlive udaje, pouze minimum:
  - sub: interni ID uzivatele
  - provider: napriklad cvut
  - providerSubject: stabilni ID z providera
  - email nebo skolni login, pokud je dostupny a potreba pro UI
  - roles: napriklad student, teacher, admin, public
  - allowedFunctions: bud ["*"], nebo seznam UDF
  - deniedFunctions: explicitni blacklist funkci
  - plan nebo licenseTier
  - iat, exp, jti
- Custom functions runtime token neoveruje kryptograficky, pokud by to vyzadovalo verejne klice a slozitou knihovnu; minimalne kontroluje expiraci a vola backend endpoint /auth/check nebo /usage/authorize podle potreby.
- Lepsi varianta pro vykon:
  - Taskpane po loginu zavola GET /auth/me.
  - Runtime cte cache evalytics.auth.entitlements.
  - UDF lokalne overi, zda je funkce v allowedFunctions.
  - Backend prubezne loguje a muze token revokovat. Revokace se projevi nejpozdeji po kratkem TTL.

#### Datovy model v MySQL

- users: id, display_name, email, primary_provider, status, created_at, updated_at, last_login_at.
- user_identities: id, user_id, provider, provider_subject, provider_username, provider_email, raw_claims_json; unikatni index (provider, provider_subject).
- roles: id, code, description.
- user_roles: user_id, role_id.
- function_policies: id, function_name nebo pattern, effect (allow/deny), role_code nebo user_id, valid_from, valid_to.
- sessions: id, user_id, refresh_token_hash, created_at, expires_at, revoked_at, user_agent_hash, last_seen_at.
- audit_events: id, user_id, session_id, event_type, function_name, result, metadata_json, created_at.

#### Backend API navrh

- GET /auth/providers: seznam provideru a zda jsou zapnute.
- GET /auth/login/{provider}/start: spusti OAuth flow.
- GET /auth/login/{provider}/callback: dokonci OAuth flow a vrati callback stranku pro Office dialog.
- POST /auth/refresh: vyda novy access/entitlement token.
- POST /auth/logout: revokuje session.
- GET /auth/me: vrati aktualniho uzivatele a opravneni.
- POST /auth/check: input { functionName }, output { allowed, reason, entitlementsVersion }.
- POST /usage/events: davkovane logovani pouziti z taskpane/runtime.

#### Integrace do UDF runtime

- Pridat novy modul nebo blok v functions.ts:
  - konstanty storage keys
  - getAuthSnapshot()
  - isTokenFresh(snapshot)
  - isFunctionAllowed(functionName, snapshot)
  - requireAuthorized(functionName)
  - unauthorized()
- Minimalni prvni verze:
  - vsechny statisticke UDF chranit
  - diagnostiky VERSION, PING nechat verejne
  - transformacni helpery STACK, UNSTACK, PARSE.NUMBER, FILL pravdepodobne nechat verejne, pokud nerozhodneme jinak
  - generatory dat mohou byt verejne nebo chranene podle produktu; rozhodnout pred implementaci
- Kazda chranena UDF zacne requireAuthorized("ANOVA").
- Pro minimalizaci rucnich zmen lze vytvorit wrapper protectedFunction(name, fn).
- Pozor na funkce s vlastnim try/catch (napr. ANCOVA/Welch diagnostiky). #UNAUTHORIZED se nesmi pohltit diagnostickym fallbackem jako bezna chyba vypoctu.

#### Integrace do taskpane UI

- Pridat v nastaveni novou cast Ucet / Account.
- Stavy: neprihlasen, prihlasovani, prihlasen, token expiroval, pristup zamitnut / ucet blokovan.
- Akce: Prihlasit pres CVUT, Odhlasit, Obnovit opravneni.
- Po uspesnem loginu ulozit auth snapshot do localStorage, zavolat full recalculation a zapsat debug log do sidepanel logu.
- V dokumentaci funkci volitelne zobrazit badge Vyzaduje prihlaseni, pokud function-docs.json dostane metadata { "auth": { "required": true } }.

#### Doporucene etapy implementace

1. Pripravit klientskou autorizacni vrstvu bez realneho OAuth: localStorage auth snapshot, requireAuthorized, unauthorized, wrapper pro chranene UDF, dev tlacitko Simulovat prihlaseni / Odhlasit.
2. Pridat metadata v function-docs.json: auth.required, pripadne auth.defaultPolicy, sidepanel badge a filtrovani podle opravneni.
3. Pridat AWS backend skeleton: API Gateway + Lambda, DB migrace, /auth/me, /auth/check, /usage/events, zatim s testovacim providerem/dev tokenem.
4. Implementovat CVUT OAuth/OIDC: overit aktualni provider dokumentaci a endpointy, nastavit client registration, redirect URI, scopes, start/callback flow, mapovani claims.
5. Zapojit realny login v taskpane: Office dialog, callback messaging, refresh/logout, recalculation po zmene stavu.
6. Zapojit server-side policy: role, blacklist, function policies, kratkodoby entitlement token, revokace po TTL.
7. Pridat audit logging: login/logout/refresh, denied UDF vzdy, allowed UDF davkovat nebo samplovat.

#### Rozhodnuti pred implementaci

- Ktere UDF maji byt verejne i bez loginu: doporuceni VERSION, PING, STACK, UNSTACK, PARSE.NUMBER, mozna FILL.
- Ktere UDF chranit: statisticke testy, ANOVA/ANCOVA, pivoty, deskriptivni statistiky podle obchodniho rozhodnuti.
- Jakou chybu presne zobrazovat: doporuceni CustomFunctions.ErrorCode.notAvailable + message #UNAUTHORIZED.
- Jak casto ma runtime kontaktovat backend: doporuceni lokalni entitlement cache + kratke TTL; backend check jen pri refreshi nebo pri chybe.
- Jak detailne logovat allowed vypocty: doporuceni denied vzdy, allowed davkovat/samplovat.
- Jak bude vypadat admin sprava prav: prvni faze muze byt prima editace DB nebo jednoduchy admin endpoint; UI az pozdeji.

#### Instrukce pro pokracovani po preruseni flow

- Nezacinej rovnou realnym CVUT OAuth. Nejprve implementuj lokalni auth vrstvu v add-inu a over chovani UDF bez backendu.
- Pred zmenami si znovu precti AI_NOTES.md, BACKLOG.md, office-addin/src/custom-functions/functions.ts, office-addin/src/public/taskpane.js, office-addin/src/public/functions.json a office-addin/src/public/function-docs.json.
- Nezapomen, ze po zmene functions.json je potreba Evalytics: Hard Reset Debug Session.
- Nepridavej secrets do repozitare. OAuth client secret, JWT signing key a DB credentials patri do AWS Secrets Manageru nebo lokalniho .env, ktery neni commitovany.
- Po kazde etape spust npm run build v office-addin/.
- Jakmile bude auth zasahovat do custom-functions metadata nebo nazvu funkci, zkontroluj tutorialy, wizard, dokumentaci CS/EN a runtime asociace.
