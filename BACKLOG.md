# Backlog - PRIORITNÍ

## Obecné

- Název pro celé řešení je Evalytics Statigo
- až bude vhodný čas, přejmenuj vše, co bude potřeba

## Web

### Cile

- Web bude mit dve role: verejny web (instalace, dokumentace, release notes) a admin UI pro spravu API dat.
- Dokumentace nesmi byt editovana natvrdo v repozitari pro kazdou zmenu. Musi existovat riditelny publika?ni flow: navrh -> revize -> schvaleni -> publikace do konkretni verze/build.
- Add-in musi umet nacitat dokumentaci podle verze buildu, aby uzivatel videl texty odpovidajici instalovanemu buildu.

### Doporucena architektura

- Frontend: `SvelteKit` (jeden projekt, route skupiny `public/*` a `admin/*`).
- Hosting webu: `S3 + CloudFront` (staticky export nebo adapter podle zvoleneho deploymentu).
- API: `API Gateway + AWS Lambda` (REST, oddelene endpointy pro public a admin).
- Databaze: `MySQL RDS` (metadatovy model pro dokumentaci, verze, workflow, audit).
- Storage dokumentacnich artefaktu: `S3` (versioned JSON snapshoty publikovanych docs).
- Cache/latence: `CloudFront` pred public docs endpointem, plus `ETag/If-None-Match`.
- Secrets: `AWS Secrets Manager` (DB credentials, signing keys, OAuth secrets).
- Monitoring: `CloudWatch Logs + Alarms`, volitelne `X-Ray`.
- v .env jsou základní přihlašovací údaje k AWS a lokální MySQL, schémata v DB jsou již bvytvořená (prázdná)

### Datovy model (minimum)

- `doc_sets`: logicky dokumentacni balik (napr. `function-docs`), jazyk, aktivni stav.
- `doc_versions`: verze docs (semantic + build stamp), stav (`draft/review/approved/published/archived`), autor, reviewer, casy.
- `doc_entries`: jednotlive funkce (`function_name`) + payload JSON pro detail.
- `doc_change_requests`: navrhy od expertu (formularove zmeny) a jejich diff.
- `doc_publications`: vazba docs verze -> publikovany artifact (`s3_key`, checksum, published_at).
- `build_doc_map`: mapovani `addin_build` -> `doc_version_id`.
- `audit_events`: kdo a kdy provedl navrh, schvaleni, publikaci, rollback.

### API kontrakty

- Public:
  - `GET /public/docs/latest?lang=cs`
  - `GET /public/docs/by-build/{build}?lang=cs`
  - `GET /public/releases`
- Admin (auth required):
  - `GET /admin/docs/versions`
  - `POST /admin/docs/versions` (create draft)
  - `GET /admin/docs/versions/{id}`
  - `POST /admin/docs/versions/{id}/changes`
  - `POST /admin/docs/versions/{id}/submit-review`
  - `POST /admin/docs/versions/{id}/approve`
  - `POST /admin/docs/versions/{id}/publish`
  - `POST /admin/docs/versions/{id}/rollback`
  - `POST /admin/build-map` (bind build -> doc version)

### Workflow pro dokumentaci

1. Expert upravi navrh pres admin formular (`draft` verze).
2. System ulozi granularni change-set (`doc_change_requests`) + audit event.
3. Reviewer schvali nebo vrati pripominky.
4. Po schvaleni se vygeneruje finalni JSON artifact (`function-docs.json` shape) a ulozi do S3.
5. Publikace vytvori zaznam v `doc_publications` a aktualizuje `build_doc_map`.
6. Add-in nacita docs podle buildu (`build_doc_map`), fallback na latest approved.
7. Pri incidentu lze provest rollback premapovanim buildu na predchozi publikaci.

### Integrace s Office add-inem

- Add-in uz ma `buildStamp`; to pouzit jako klic pri fetch docs.
- Fetch endpoint preferovat: `/public/docs/by-build/{build}?lang=...`.
- Fallback poradi: by-build -> latest-approved -> lokalni bundled JSON.
- Pri chybach neblokovat funkce; zobrazit degradovany docs stav a logovat diagnostiku.

### Bezpecnost a role

- Role minimalne: `viewer`, `editor`, `reviewer`, `publisher`, `admin`.
- `publish` a `build-map` jen pro `publisher/admin`.
- Vsechny admin akce auditovat (kdo, co, kdy, pred/po).
- Validovat payload proti JSON schema (per function docs shape).

### CI/CD a prostredi

- Prostredi: `dev`, `staging`, `prod` oddelene (S3 bucket/API stage/RDS schema).
- Pipeline kroky:
  1. lint + typecheck
  2. unit/integration testy API
  3. schema migration (staging/prod guard)
  4. deploy Lambd
  5. deploy webu
  6. smoke test public docs endpointu
- Zakaz primeho publish do prod bez schvaleni (manual gate).

### Implementacni plan po etapach

1. Foundation
- Zalozit databazovy model, migrace, skeleton Lambda API, health endpointy.
- Pridat `/public/docs/latest` a nacitani statickeho JSON ze S3.

2. Versioned docs
- Implementovat `doc_versions` + `doc_entries` + generate/publish artifact do S3.
- Pridat `/public/docs/by-build/{build}` + fallback logiku.

3. Admin UI MVP
- Formulare pro editaci funkci, seznam verzi, diff preview, submit-review.
- Role-based access (zatim muze bez finalniho CVUT OAuth, ale pripraveno na napojeni).

4. Review/publish governance
- Approve/publish/rollback endpointy, audit eventy, build mapping.
- Guardrails proti publish bez reviewer approve.

5. Add-in integration
- Upravit taskpane fetch docs na build-aware endpoint.
- Pridat robustni fallback a telemetry pro docs load failures.

6. Hardening
- Caching policy, ETag, rate limiting, alarms, backup/restore runbook.

### Rozhodnuti pred implementaci

- Zdroj pravdy pro docs: DB-first + S3 artifact (doporuceno) nebo stale repo-first.
- Granularita mapovani: build->exact docs verze (doporuceno) vs latest docs.
- Schvalovaci model: one-step approve/publish nebo two-man rule pro produkci.
- Kde bude auth provider pro admin web v prvni fazi (CVUT OAuth hned vs interim login).

### Instrukce pro dalsi AI session

- Pred praci nacist: `BACKLOG.md`, `AI_NOTES.md`, `website/`, `office-addin/src/public/taskpane.js`.
- Nejdriv implementovat etapu 1 a 2; nepreskakovat rovnou na admin UI publish flow.
- Pri kazde etape zapsat do backlogu: co je hotove, co chybi, jaka jsou rozhodnuti.
- Nemodifikovat produkcni credentials v repu; pouzivat pouze env/secrets manager.

### Stav implementace (2026-05-08)

- Hotovo:
  - `website/` bylo resetovano na novy foundation skeleton (public + admin shell).
  - Pridana API kostra v `api/` s handlery pro `health`, `public/docs/latest`, `public/docs/by-build/{build}`, `admin/docs/versions`, `admin/build-map`.
  - Pridan OpenAPI navrh v `api/openapi.yaml`.
  - Pridana SQL baseline migrace v `migrations/001_web_docs_foundation.sql`.
  - Pridany infra poznamky pro nasazeni na `statigo.evalytics.org` v `infra/README.md`.
  - Etapa 2 backend wiring: public docs endpointy jsou napojene na `build_doc_map` / `doc_publications` lookup v MySQL a nacitani artifactu ze S3.
  - `public/docs/by-build/{build}` implementuje fallback na latest published docs, pokud mapovani build->verze neexistuje.
  - Admin endpoint `admin/docs/versions` ma zakladni bearer auth gate (`ADMIN_BEARER_TOKENS`) a read endpoint vraci latest publikace pro `cs/en`.
  - Admin endpoint `POST /admin/docs/versions` je implementovan pro realne vytvoreni draft verze (`code`, `language`, `versionLabel`, volitelne `addinBuild`).
  - Admin endpoint `POST /admin/build-map` je implementovan pro svazani `addinBuild -> docVersionId` s automatickym uzavrenim predchozi aktivni mapy.
  - Pridany bootstrap podklady pro runtime:
    - `api/.env.example`
    - `scripts/bootstrap-api-env.ps1` (mapuje root `.env` -> `api/.env.local`, nove umi `-Environment dev|prod`)
    - `migrations/002_web_docs_seed_minimal.sql` (minimalni seed)
  - AWS runtime priprava:
    - vytvoreny bucket `statigo-evalytics-org-docs-dev`
    - vytvoreny bucket `statigo-evalytics-org-docs-prod`
    - na obou bucketech zapnuto `Block Public Access`, `Versioning`, `SSE-S3 (AES256)`
  - Auth MVP (tag-based):
    - migrace `003_auth_tag_model.sql` (`auth_users`, `auth_tags`, `auth_user_tags`, `auth_function_policies`)
    - endpointy `POST /auth/dev-login`, `GET /auth/me`, `POST /auth/check`
    - admin endpointy `GET /admin/auth/users`, `POST /admin/auth/users/{userId}/tags`, `GET/POST /admin/auth/tags`
    - web admin stranka `/admin` umi nacist users/tagy a prirazovat tagy uzivatelum
    - custom functions runtime umi fallbacknout na vzdaleny `auth/check` (Bearer + API base URL v auth snapshotu)
  - Infra IaC pro produkcni domenu:
    - `infra/web-static.yaml` pro `statigo.evalytics.org` (S3 + CloudFront + Route53)
    - `infra/sam-api.yaml` pro `api.statigo.evalytics.org` (HTTP API + Lambda + custom domain + Route53)
    - `scripts/deploy-infra-prod.ps1` pro jednotny deploy flow
  - Produkcni deploy (2026-05-09):
    - ACM certifikaty vydany:
      - `statigo.evalytics.org` (us-east-1)
      - `api.statigo.evalytics.org` (eu-central-1)
    - nasazen stack `statigo-web-prod` (CloudFront + S3 + Route53)
    - nasazen stack `statigo-api-prod` (HTTP API + Lambda + custom domain + Route53)
    - smoke check:
      - `https://api.statigo.evalytics.org/health` vraci `status=ok`
      - `https://statigo.evalytics.org` vraci HTML
- Chybi:
  - Endpoints pro changes/review/approve/publish/rollback.
  - Jemnejsi RBAC (role model) misto jednoducheho shared bearer token seznamu pro admin endpointy.
  - JSON schema validace docs payloadu a diff workflow.
  - Seed publikovanych docs (`doc_publications` + validni S3 artifacty) pro plne funkcni public endpointy v dev/prod.
  - CI/CD deploy pipeline a smoke testy.
  - Napojeni realneho OAuth providera (CVUT) misto `auth/dev-login`.
  - Doresit publish flow docs (`doc_publications` + artifact upload + endpoint smoke test pro latest/by-build).
- Rozhodnuti:
  - Pokracovat etapou 3: navazat endpointy changes/review/approve/publish + doplnit admin UI MVP.

## Nefunkční UDF

- aktuálně v pořádku

## UI

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

## Regrese (implementace)

- [x] Pridana nova UDF `REGRESSION` (linearni OLS report: model summary, ANOVA, koeficienty, diagnostika).
- [x] Pridana nova UDF `REGRESSION.PREDICT` (bodova predikce + CI pro stredni hodnotu).
- [x] Pridany wizardy pro obe funkce (argumenty jsou konfigurovatelne ve standardnim wizard panelu).
- [x] Pridana demo data a demo formule pro obe funkce v sidepanelu.
- [x] Korelacni metody `CORREL.SPEARMAN` a `CORREL.MATRIX` presunuty v UI katalogu do sekce `Regrese/Regression`.
- [ ] Dodelat bohatsi lokalizovanou dokumentaci `function-docs.json` pro nove regresni funkce (CZ/EN popisy parametru a vystupu).
- [ ] Rozsirit `REGRESSION.PREDICT` o predikcni interval (PI), nejen CI.

## Regrese (2.1 + 2.2) - aktualizace

- [x] Doplnena `REGRESSION.SELECT` (forward/backward/stepwise; kriterium adjR2/AIC/BIC; trace kroku).
- [x] Doplnena `TREND.FIT` (linear/log/exp/power).
- [x] Doplnena `TREND.COMPARE` (vice trend modelu v jednom vystupu).
- [x] Doplnena demo data + demo formule pro `REGRESSION.SELECT`, `TREND.FIT`, `TREND.COMPARE`.
- [x] Doplnen wizard setup pro vsechny nove metody v sidepanelu.
- [x] Dokumentace sjednocena v jednom souboru `office-addin/src/public/function-docs.json` (CZ/EN) pro vsechny nove metody.

## Reorganizace sekce Testy + nove neparametricke metody

- [x] UI sekce `Testy/Tests` zrusena na urovni katalogu a nahrazena dvojici:
  - `Srovnání skupin / Group Comparisons`
  - `Kvalitativní data / Categorical Data`
- [x] `CONTINGENCY.*` presunuto do `Kvalitativní data / Categorical Data`.
- [x] Doplnena UDF `KRUSKAL.WALLIS` (Kruskal-Wallis ANOVA).
- [x] Doplnena UDF `FRIEDMAN.ANOVA` (Friedmanova ANOVA).
- [x] Doplnena UDF `JONCKHEERE.TERPSTRA` (Jonckheere-Terpstruv test trendu).
- [x] Ke kazde nove metode dodelan wizard + demo v sidepanelu.
- [x] Ke kazde nove metode dodelana CZ/EN dokumentace v `office-addin/src/public/function-docs.json`.
