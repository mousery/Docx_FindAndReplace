# Docx_FindAndReplace

You know... find and replace, from microsoft document. Just trying to replicate it with python-docx.

Paragraphs in Microsoft document are actually split into sections of text-containing objects called runs, and find and replacing text in paragraphs may not be as intuitive as it sounds. So I wrote find_and_replace to make it easier at least for me.

It is also regular expression compatible, and can use re syntax e.g. (?P<name>something).

## Example

    insert images here later

## Others
1. get_font_name and set_font_name (including chinese, arabics and others font name)

    For some reason python-docx don't support obtaining chinese, arabics and other font name of text-containing object unless going down to the _element level. I wrote get_font_name, change_font and change_font_from_change_dict to make it easier at least for me.

2. set_run_text
3. add_run_after_run
4. remove_run
5. split_run_at_string_index
6. delete_paragraph
7. isolate_para_runs_by_span