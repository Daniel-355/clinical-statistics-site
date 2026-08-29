local function text(meta, key)
  if meta[key] == nil then return nil end
  return pandoc.utils.stringify(meta[key])
end

local function esc(value)
  return value:gsub("&", "&amp;"):gsub("<", "&lt;"):gsub('"', "&quot;")
end

local function list_items(value)
  if value == nil then return {} end
  if value.t == "MetaList" then return value end
  return { value }
end

function Pandoc(doc)
  local section = text(doc.meta, "section")
  local category = text(doc.meta, "category")
  local topic = text(doc.meta, "topic")
  if not section or not category or not topic then return doc end

  local nav = {
    ["Clinical Trial Process & Design"] = "clinical-trial-process-design",
    ["Data Management & Programming"] = "data-management-programming",
    ["Biostatistics"] = "biostatistics",
    ["Resources"] = "resources",
    ["Epidemiology & Real-World Evidence"] = "epidemiology-rwe",
    ["About"] = "about"
  }
  local section_slug = nav[section] or ""
  local category_slug = category:lower():gsub("[^%w]+", "-"):gsub("^-", ""):gsub("-$", "")
  local breadcrumb = '<nav class="breadcrumbs" aria-label="Breadcrumb"><a href="../../index.html">Home</a> <span aria-hidden="true">›</span> <a href="../../_generated/' .. section_slug .. '/index.html">' .. esc(section) .. '</a> <span aria-hidden="true">›</span> <a href="../../_generated/' .. section_slug .. '/index.html#' .. category_slug .. '">' .. esc(category) .. '</a> <span aria-hidden="true">›</span> <span aria-current="page">' .. esc(topic) .. '</span></nav>'
  local post_type = text(doc.meta, "type") or "article"
  local header = '<div class="post-header-card">' .. breadcrumb .. '<div class="post-meta">' .. esc(section) .. ' · ' .. esc(category) .. ' · ' .. esc(post_type) .. '</div></div>'
  table.insert(doc.blocks, 1, pandoc.RawBlock("html", header))

  local tail = {}
  local project_dir = os.getenv("QUARTO_PROJECT_DIR") or "."
  local ok, manifest = pcall(dofile, project_dir .. "/_generated/post-media.lua")
  local media = ok and manifest[text(doc.meta, "title") or ""] or nil
  local videos = media and media.youtube or list_items(doc.meta.youtube)
  for _, video in ipairs(videos) do
    local id = pandoc.utils.stringify(video)
    if id ~= "" then
      table.insert(tail, pandoc.RawBlock("html", '<section class="video-panel"><h2>Video</h2><div class="video-wrap"><iframe src="https://www.youtube-nocookie.com/embed/' .. esc(id) .. '" title="YouTube video" loading="lazy" allowfullscreen></iframe></div></section>'))
    end
  end

  local pdfs = media and media.pdf or list_items(doc.meta.pdf)
  for _, item in ipairs(pdfs) do
    local title = item.title and pandoc.utils.stringify(item.title) or "PDF"
    local file = item.file and pandoc.utils.stringify(item.file) or ""
    if file ~= "" then
      local html = '<section class="pdf-panel"><h2>' .. esc(title) .. '</h2><object class="pdf-frame" data="' .. esc(file) .. '" type="application/pdf"><p>Preview unavailable. <a href="' .. esc(file) .. '">Open the PDF</a>.</p></object><p><a href="' .. esc(file) .. '" target="_blank">Open PDF</a> · <a href="' .. esc(file) .. '" download>Download PDF</a></p></section>'
      table.insert(tail, pandoc.RawBlock("html", html))
    end
  end

  local downloads = media and media.downloads or list_items(doc.meta.downloads)
  if #downloads > 0 then
    local rows = {}
    for _, item in ipairs(downloads) do
      local title = item.title and pandoc.utils.stringify(item.title) or "Download"
      local file = item.file and pandoc.utils.stringify(item.file) or ""
      if file ~= "" then table.insert(rows, '<div class="download-item"><span>' .. esc(title) .. '</span><a href="' .. esc(file) .. '" download>Download</a></div>') end
    end
    if #rows > 0 then table.insert(tail, pandoc.RawBlock("html", '<section class="downloads-panel"><h2>Downloads</h2>' .. table.concat(rows) .. '</section>')) end
  end
  for _, block in ipairs(tail) do table.insert(doc.blocks, block) end
  return doc
end
