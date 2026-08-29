# Clinical Statistics Notes

A static Quarto knowledge base for clinical research, data management, biostatistics, resources, and real-world evidence. Content and configuration live in Git; production needs no server, database, CMS, or JavaScript application.

## The files you normally edit

- `posts/<post-name>/index.qmd` to write or update a post.
- `posts/<post-name>/files/` and `images/` for that post's attachments.
- `config/navigation.yml` to change Level 1 or Level 2 navigation.
- `styles.css` to change the site's visual style.
- `_quarto.yml` and everything in `_generated/` are generated. Do not edit them; change navigation, post metadata, or the build script instead.

## Add a post

From R:

```r
source("scripts/new_post.R")
new_post(
  title = "MMRM in Clinical Trials",
  section = "Biostatistics",
  category = "Longitudinal & Repeated Measures",
  topic = "MMRM",
  type = "tutorial"
)
```

Or from a terminal:

```bash
Rscript scripts/new_post.R "MMRM in Clinical Trials" "Biostatistics" "Longitudinal & Repeated Measures" "MMRM" tutorial draft
```

Write in the new `posts/mmrm-in-clinical-trials/index.qmd`. Posts start as drafts. Change `status: draft` to `status: published` when ready, then regenerate and validate.

## Change navigation and topics

- **Add a Level 1 menu:** add one item with `title`, `slug`, and `children` in `config/navigation.yml`.
- **Add a Level 2 menu:** add one child title under the appropriate Level 1 item in `config/navigation.yml`.
- **Add a Level 3 topic:** no navigation edit is needed. Set `topic:` in a post. Published posts are grouped by topic automatically.
- **Change a post category:** edit its `section:` and/or `category:` metadata. Values must match `config/navigation.yml` exactly.
- **Move a post:** normally change metadata only. The folder may also be renamed, but every post must keep its own `index.qmd`, `files/`, and `images/` together.

After metadata or navigation changes:

```bash
Rscript scripts/build_navigation.R
Rscript scripts/validate_site.R
```

## Attachments, PDFs, and video

Put every attachment in the post's `files/` folder.

Add a normal download:

```yaml
downloads:
  - title: "Example dataset"
    file: "files/example.csv"
```

Add a PDF viewer and, if desired, the same PDF as a download:

```yaml
pdf:
  - title: "Lecture notes"
    file: "files/lecture.pdf"
downloads:
  - title: "Lecture notes"
    file: "files/lecture.pdf"
```

Embed YouTube using only the video ID:

```yaml
youtube:
  - "VIDEO_ID"
```

The shared post filter creates responsive video, PDF, breadcrumb, metadata, and download sections automatically. Empty sections are omitted.

## Resource-only, featured, and draft posts

- Set `type: resource` for a short description plus attachments.
- Set `featured: true` to place a published post on Home automatically.
- Set `status: draft` to exclude a post from production, search, and all listings.
- Set `status: published` to publish it after validation.

All three post types—`article`, `tutorial`, and `resource`—use the same metadata and layout.

## Preview and validate

Install [Quarto](https://quarto.org/docs/get-started/) once, then run:

```bash
Rscript scripts/build_navigation.R
Rscript scripts/validate_site.R
quarto preview
```

To preview a draft directly without putting it in production listings:

```bash
quarto preview posts/draft-missing-data-note/index.qmd
```

The validator checks metadata, section/category values, slugs, attachment paths, PDF paths, and draft leakage. The render also checks that draft HTML was not emitted.

## Publish with GitHub Pages

1. Push the repository to GitHub with `main` as the default branch.
2. In repository **Settings → Pages**, choose **GitHub Actions** as the source.
3. Push a commit. `.github/workflows/publish.yml` installs R, YAML, and Quarto; generates; validates; renders; and deploys `_site`.
4. The configured production URL is `https://daniel-355.github.io/clinical-statistics-site`.

For a custom domain, add a `CNAME` file containing the domain, add the matching DNS record at the domain provider, and configure the domain under GitHub Pages settings. GitHub Pages supplies HTTPS after DNS validation.

## Restore an older version

Inspect history with `git log --oneline`. The safest recovery is to revert a change with a new commit:

```bash
git revert <commit-id>
git push
```

To recover one file without rewriting history:

```bash
git restore --source <commit-id> -- path/to/file
git commit -m "Restore earlier file version"
git push
```

## Ask Codex for maintenance

Useful requests are deliberately small: “create a draft tutorial under Biostatistics → Time-to-Event Analysis → Cox Regression,” “add this PDF to the Kaplan-Meier post,” or “add Diagnostic Accuracy under Biostatistics and regenerate.” Codex should edit source configuration/content, run generation and validation, and never hand-edit `_generated/`.

## Dependencies

- Quarto CLI
- R (only at build time)
- R package `yaml`
- Git and GitHub Pages

The rendered site uses Quarto's built-in Bootstrap theme, search, code highlighting, math, cross-references, and minimal native browser media features. There is no Node backend, database, PHP, or CMS.

## Known Version 1 limits

- Related-post recommendations are represented by section/topic browsing rather than a separate ranking system.
- Drafts are previewed directly by file; they are intentionally absent from the full-site preview.
- PDF display depends on the visitor's browser; an open/download fallback is always shown.
