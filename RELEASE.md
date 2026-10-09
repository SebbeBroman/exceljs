# Publishing 0.2.0

The package version is already `0.2.0`. The changelog and migration guide cover
CSV/protection API changes, compact loaded validation ranges, legacy removal,
loaded-table fixes, bundle reductions and regression checks.

From the repository root:

```sh
pnpm install --frozen-lockfile
pnpm release:pack
```

This runs lint/format, all tests, source/spec/public type checks, a clean build,
Node/ESM and browser runtime checks, dependency/size budgets, then packs locally.
The tarball and its checked file metadata are saved under `build/release/`.
No registry write occurs. `pnpm release:check` runs just the validation.

When ready to publish, sign in to npm with an account authorized for
`@sebbebroman/exceljs`, then run:

```sh
pnpm release
```

The publish hook repeats `release:check`; npm may request your publishing OTP.
The package publishes publicly to npm as `0.2.0` with the default `latest` tag.
After publication, verify the registry version and tag the release commit:

```sh
npm view @sebbebroman/exceljs version
# From the commit that was published:
git tag v0.2.0
git push origin master v0.2.0
```

Registry availability and npm authorization are checked by npm at publish time.
The pinned dependency `@sebbebroman/fast-csv@0.1.0` must remain available.
