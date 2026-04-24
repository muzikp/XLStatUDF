# AWS Deployment

This website is prepared for static deployment to AWS with:

- `S3` for built website files and downloadable installers
- `CloudFront` for HTTPS, caching, and global delivery
- `Route 53` for the public DNS record such as `xlstat.evalytics.org`

## Recommended target

- hostname: `xlstat.evalytics.org`
- public URL: `https://xlstat.evalytics.org`

## Suggested architecture

1. Create an S3 bucket for the site content.
2. Keep the bucket private.
3. Create a CloudFront distribution with Origin Access Control to the bucket.
4. Request an ACM certificate for `xlstat.evalytics.org` in `us-east-1`.
5. Attach the certificate to CloudFront.
6. Add a Route 53 alias record for `xlstat.evalytics.org` to the CloudFront distribution.

## Required environment variables

The deployment script reads values from the repository root `.env` by default.

- `AWS_REGION`
- `AWS_S3_BUCKET`
- `AWS_CLOUDFRONT_DISTRIBUTION_ID`
- `AWS_ACCESS_KEY_ID`
- `AWS_SECRET_ACCESS_KEY`

Optional:

- `AWS_SESSION_TOKEN`
- `AWS_HOSTNAME`
- `AWS_SITE_URL`

See [`.env.example`](./.env.example) for the expected keys.

## Deploy command

From the `website/` directory:

```powershell
npm install
npm run deploy:aws
```

This will:

1. load variables from `../.env`
2. bump the deployment version in `../version.json`
3. rebuild the Excel add-in and installers with the new version
4. build the website
5. upload static files to S3
6. invalidate CloudFront

## Notes

- HTML files are uploaded with a short cache policy.
- versioned assets and downloads are uploaded with a long immutable cache policy.
- the website build already includes the synchronized documentation and copied installer files.
- deployment versioning currently follows `1.0.x`, where `x` increments by one on every publish.
- ACM certificates for CloudFront must be created in `us-east-1`, even if the bucket is elsewhere.
