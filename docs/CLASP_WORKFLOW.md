# CLASP Workflow

Bound Apps Script project:
- `2026_WASP-Katana-In`
- Script ID: `1vMGYLeN5iCcLzW9mF0DJB3bT-My_yj1Tv8MVVnaVWKLb4m582ZqyT4Lg`

Local source of truth:
- [src](C:\Users\demch\OneDrive\Desktop\final-katana-wasp-project\DELETE-ME-LATER\src)

Run from:
- `C:\Users\demch\OneDrive\Desktop\final-katana-wasp-project\DELETE-ME-LATER`

Common commands:

```powershell
npm run clasp:whoami
npm run clasp:status
npm run clasp:pull
npm run clasp:push
npm run clasp:open
npm run clasp:versions
npm run clasp:deployments
```

Create a version:

```powershell
npm run clasp:version -- "activity renderer update"
```

Create a new deployment:

```powershell
npm run clasp:deploy -- --description "Codex deploy"
```

Redeploy an existing deployment:

```powershell
npm run clasp:deploy -- --deploymentId <DEPLOYMENT_ID> --description "Codex redeploy"
```

Notes:
- `src/.clasp.json` is the active local Apps Script binding.
- `src/.claspignore` allows only `.js` files and `appsscript.json` to be pushed.
- `HANDOFF_AGENT.txt`, `.clasp.json`, and other local files are not pushed.
