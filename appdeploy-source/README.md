# Hill School AI Inventory — AppDeploy source snapshot

This directory contains the source snapshot exported from the existing AppDeploy application.

- React/Vite frontend: `src/`
- AppDeploy backend handlers: `backend/`
- Build configuration and test definitions are included.

## Deployment note
This is a source archive, not a standalone GitHub Pages deployment. The frontend imports `@appdeploy/client` and the backend imports `@appdeploy/sdk`; these are platform-specific integrations. GitHub Pages does not run backend handlers or provide the existing database/authentication services. To run this independently from AppDeploy, these integrations must be migrated to a supported backend and database/authentication must be configured separately.

No live database records or secrets are included in this source snapshot.
