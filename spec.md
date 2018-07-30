## Building & Deployment

Lambda functions exist in different regions.
Deployment happens automatically, depending on environment.

- Dev branch will automatically build & deploy to ews.yasoon.org
  - Build new version
  - Upload function code to lambda
  - Create new version (auto-generate)
  - Switch "dev" alias to new version
- Master branch will deploy to ews.yasoon.com
  - On push, new version will be build and uploaded to lambda
  - new version will be created (auto-generate)
  - Alias of prod needs to be changed manually! 