---
# create lunr store for search page
---
{%- assign items = site.data[site.metadata] -%}
{%- assign fields = site.data.config-search -%}
var page_info = [
{%- for page in site.pages -%}
{"id": {{page.permalink | jsonify}}, "title": {{page.title | jsonify}}, "content": {{page.content | strip_html | jsonify }}, "excerpt": {{page.content | strip_html | truncatewords: 20 | jsonify | remove: "#" }}}{%- unless forloop.last -%},{%- endunless -%}
{%- endfor -%}
];
var metadata_info = [
{%- for item in items -%} 
{ "id": {{ item.objectid | jsonify }}, {% for f in fields %}{{ f.field | jsonify }}: {% if item[f.field] %}{{ item[f.field] | normalize_whitespace | replace: '""','"' | jsonify }}{% else %}"none"{% endif %}{% unless forloop.last %},{% endunless %}{% endfor %} }{%- unless forloop.last -%},{%- endunless -%}
{%- endfor -%}
];

var store = metadata_info.concat(page_info);