# Website Build and Maintenance Guide

This project is a static clinical statistics knowledge website built with **Quarto, R, and GitHub Pages**. It does not use a database, WordPress, React, or a server backend. Its primary goals are simplicity, stability, low cost, and long-term maintainability.

## Core Structure

- `config/navigation.yml`: The single source of truth for Level 1 navigation and Level 2 categories.
- `posts/<slug>/index.qmd`: Article content and metadata.
- `posts/<slug>/files/`: PDFs, code, datasets, and other downloadable files belonging to a post.
- `posts/<slug>/images/`: Images belonging to a post.
- `styles.css`: Site-wide styling.
- `scripts/build_navigation.R`: Generates navigation, section pages, homepage listings, and `_quarto.yml`.
- `scripts/validate_site.R`: Validates post metadata, categories, and attachment paths.
- `_generated/`: Disposable generated files. Do not edit them manually.
- `.github/workflows/publish.yml`: Automated GitHub Pages publishing workflow.

## Create a New Post

Run the following in R:

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

Then edit `index.qmd` in the newly created post directory. New posts default to `status: draft`. Change the status when the post is ready to publish:

```yaml
status: published
```

## Change Navigation

- To add a Level 1 menu, edit `config/navigation.yml` and add an item containing `title`, `slug`, and `children`.
- To add a Level 2 category, add its title under the appropriate Level 1 item's `children` list.
- To add a Level 3 topic, set `topic:` in a post's metadata. Topic groups are created automatically; no navigation edit is required.

Regenerate the site configuration after changing navigation or post metadata:

```bash
Rscript scripts/build_navigation.R
```

## Add Attachments

Place each attachment in the post's `files/` directory, then declare it in the post YAML:

```yaml
downloads:
  - title: "Example Dataset"
    file: "files/example.csv"
```

To display a PDF in the browser:

```yaml
pdf:
  - title: "Lecture Notes"
    file: "files/lecture.pdf"
```

To embed a YouTube video, provide only its video ID:

```yaml
youtube:
  - "VIDEO_ID"
```

## Validate and Preview Locally

```bash
Rscript scripts/build_navigation.R
Rscript scripts/validate_site.R
quarto preview
```

To perform a complete production build:

```bash
quarto render
```

## Publish the Website

Commit and push the changes to GitHub:

```bash
git add .
git commit -m "Update website content"
git push
```

GitHub Actions automatically regenerates the site, validates it, renders the Quarto project, and deploys the result to GitHub Pages.

Production website:

<https://daniel-355.github.io/clinical-statistics-site/>

## Maintenance Rules

Files normally edited during routine maintenance:

- `posts/*/index.qmd`
- Each post's `files/` and `images/` directories
- `config/navigation.yml`
- `styles.css`

Do not manually edit:

- `_generated/`
- `_quarto.yml`
- `_site/`

If generated content is incorrect, fix the navigation configuration, post metadata, or R build logic, and then regenerate the site.
