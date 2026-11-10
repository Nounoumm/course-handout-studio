# Plugin market index

`index.json` is a [DSH plugin market index](https://github.com/deepseek-ai/deepseek-harness) in **format 1**. The Web **Discover** tab and the read-only `plugin_market` tool read it; installation still goes through `plugin_manager`.

Public URL:

```text
https://raw.githubusercontent.com/Nounoumm/course-handout-studio/main/market/index.json
```

## Add it to Discover

1. Open **Plugins → More** (Experimental plugins) and enable **Custom plugin markets**.
2. Open **Discover → Add index** and paste the URL above.
3. Install **Course Handout Studio** from the list.

DSH has no official central market; an index is a normal HTTPS-hosted JSON document that each user opts into. This repository's index is the authoritative listing for this plugin.

## What the entry points to

The single version `1.0.0` installs from Git, pinned to a commit:

```text
git+https://github.com/Nounoumm/course-handout-studio.git#7808c128cc3f70230ee1ca2b4bad288fb0e32695
```

Pinning keeps an installed version reproducible even if `main` moves. The pinned commit must contain the plugin package at the repository root.

## Releasing a new version

1. Update `version` in `package.json` and commit the code.
2. Tag the commit, e.g. `v1.0.1`, and push the tag.
3. In `index.json`, add a version object with the new SemVer, the commit SHA, and `publishTimestamp` (Unix seconds).
4. Keep the old version object; the reader sorts versions by SemVer and defaults to the highest compatible one.
5. Validate before pushing:

```js
import { parseMarketIndex, marketInstallTarget } from '@deepseek-ai/dsh-plugin-market-index'
import { readFileSync } from 'node:fs'

const index = parseMarketIndex(
  JSON.parse(readFileSync('market/index.json', 'utf8')),
  'https://raw.githubusercontent.com/Nounoumm/course-handout-studio/main/market/index.json',
)
for (const plugin of index.plugins) {
  for (const version of plugin.versions) {
    console.log(plugin.name, version.version, marketInstallTarget(plugin.name, version))
  }
}
```

## Format notes

- `title` and every `description` must be either a string or an object containing an `en` key; other language keys use lowercase language ids.
- `metadata.categories` and `metadata.labels` declare the ids the plugin entry may reference.
- A Git source requires a canonical `https://` repository URL with at least two path segments and a lowercase 40-hex commit SHA.
- `publishTimestamp` is an integer number of Unix seconds.
- Relative icons resolve against the index URL; this index uses an absolute `https://` icon.
