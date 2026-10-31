# How to publish xlsx-audit

Publishing runs from GitHub Actions with PyPI's trusted publishing, so there is no PyPI token to create, store or leak. The owner does the account steps once; after that, a GitHub release publishes a new version.

## Before the first release

1. **Try it on real workbooks.** The tests use workbooks built from raw XML. Run `xlsx-audit lint` and `xlsx-audit diff` on two or three real models saved by Excel, and note anything that looks wrong or noisy.
2. **Decide whether the GitHub repo goes public.** The package is MIT-licensed, and the PyPI page links to the repo, so links stay broken while it's private. The pre-commit hook in the README also needs a public repo.

## One-time setup (owner)

1. **PyPI account** at <https://pypi.org/account/register/>, with two-factor authentication turned on.
2. **Pending trusted publisher.** On PyPI: *Your account → Publishing → Add a new pending publisher → GitHub*, with:

   | Field | Value |
   |---|---|
   | PyPI Project Name | `xlsx-audit` |
   | Owner | `gravitybendr` |
   | Repository name | `xlsx-audit` |
   | Workflow name | `publish.yml` |
   | Environment name | `pypi` |

   A pending publisher doesn't reserve the name; someone else could still register `xlsx-audit` first. Publish soon after adding it.
3. **GitHub environment.** In the repo: *Settings → Environments → New environment*, named `pypi`. Optionally add yourself under *Required reviewers*, so each release waits for your approval before it reaches PyPI.

## Each release

1. Update `version` in `pyproject.toml` and `__version__` in `src/xlsx_audit/__init__.py`, and add a `CHANGELOG.md` entry.
2. Push to `main` and wait for the CI workflow to pass.
3. Create a GitHub release with the tag `v` + the version, for example:

   ```
   gh release create v0.1.0 --title "v0.1.0" --notes-file CHANGELOG.md
   ```

4. The **Publish to PyPI** workflow tests, builds and uploads. It stops if the tag doesn't match the version in `pyproject.toml`.
5. Check <https://pypi.org/project/xlsx-audit/>, then `pip install -U xlsx-audit` in a fresh environment.

## Listing

- **PyPI page:** built from `README.md`, plus the summary, keywords and classifiers in `pyproject.toml`, **as they were in the release**. 0.1.0 is live with the old README; the new README, summary and keywords appear on PyPI with the next release.
- **README images** load from `https://raw.githubusercontent.com/gravitybendr/xlsx-audit/main/docs/...`. They show on GitHub and PyPI only after `docs/` is pushed to `main`, so push before releasing. (The repo is public; if it ever goes private, the images stop loading on PyPI.)
- **Listing copy:** `marketing/pypi.md` (summary, image captions, keywords) and `marketing/github.md` (repo description, topics, social preview, and the `gh repo edit` command to apply them).
- **Search phrases worth a mention when sharing:** "excel formula audit", "xlsx lint", "spreadsheet diff", "financial model review", "find hardcoded values in excel", "excel #REF! finder".

## Images

`scripts/make_images.py` makes every image from real output:

```
.venv\Scripts\python.exe scripts\make_images.py
```

It rebuilds the demo workbooks in `docs/demo/` (`scripts/demo_workbooks.py`), runs the CLI on them and saves the output to `docs/demo/output/`, writes HTML pages to `build/images/`, and renders them with headless Chromium in Docker (`scripts/render_pages.sh`, image `mcr.microsoft.com/playwright:v1.64.1-noble`). It writes `docs/images/*.png` and the logo PNGs and social preview in `docs/brand/`. The logo source is `docs/brand/logo.svg`.

When the output changes (a new version, a changed message), run it again and copy the new text from `docs/demo/output/lint.txt` and `diff.txt` into the README examples.

## Paid tier (not built yet)

The plan is a free open-source core on PyPI and paid extras sold through Polar (merchant of record, licence keys). Candidates: an HTML review report, report history across versions, custom rule packs for a team's modelling standards, and batch scanning of folders. Polar pays out in about 120 countries, but not to mainland China; confirm the owner's payout country before building it.
