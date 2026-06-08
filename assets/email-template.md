---
tags:
  - email
  - {{folder}}
date: {{date}}
from: "{{sender}}"
to: {{to}}
subject: "{{subject}}"
---

# {{subject}}

**From:** {{sender_name}} <{{sender}}>
**To:** {{to_list}}
{{#if cc}}**CC:** {{cc_list}}{{/if}}
**Date:** {{date_formatted}}
{{#if has_attachments}}**Attachments:** {{attachment_list}}{{/if}}

---

{{body}}

{{#if attachments}}
## Attachments

{{#each attachments}}
- {{name}} ({{size_formatted}})
{{/each}}
{{/if}}

<!-- 
  NOTE: This is a reference template only — there is currently no built-in
  template renderer in the CLI. The export service hardcodes its markdown
  format. This template is provided as a guide for the expected structure
  of exported email files.
-->
