# Journal profiles

One folder per journal template, made by the skill the first time a template is used
(`make_docx.ps1 -TemplatePath …` exits 2 and prepares the evidence; see
`references/new_journal_profile.md`). Later runs with the same template file find the profile by
its SHA-256.

A profile contains the journal's own template (`template_reference.docx`). Check the journal's
terms before publishing one; profiles are not part of the public repository by default.
