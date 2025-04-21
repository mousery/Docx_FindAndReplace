# Docx_FindAndReplace

You know... find and replace, from microsoft document. Just trying to replicate it with python-docx.

Paragraphs in Microsoft document are actually split into sections of text-containing objects called runs, and find and replacing text in paragraphs may not be as intuitive as it sounds. So I wrote find_and_replace to make it easier at least for me.

It is also regular expression compatible, and can use re syntax e.g. (?P<name>something).

## Example 1

### test1.docx -> output1.docx
<img src="https://github.com/mousery/Docx_FindAndReplace/raw/main/example/example-test1.docx.png" height="300">  <img src="https://github.com/mousery/Docx_FindAndReplace/raw/main/example/example-output1.docx.png" height="300">

    from docx import Document as DOCX
    d1 = DOCX("./example/test1.docx")
    find_and_replace(d1.paragraphs[0], "([a-zA-Z]+)([0-9]+)(?P<r>[a-zA-Z]+)", r"\g<r>\2\1\3")
    find_and_replace(d1.paragraphs[1], "(!+)([\[\]]+)(?P<v>v{4})(v+)", r"abcd\g<v>\1\g<v>\2\g<v>")
    d1.save("./example/output1.docx")

## Example 2

### test2.docx -> output2.docx
<img src="https://github.com/mousery/Docx_FindAndReplace/raw/main/example/example-test2.docx.png" height="300">  <img src="https://github.com/mousery/Docx_FindAndReplace/raw/main/example/example-output2.docx.png" height="300">

    from docx import Document as DOCX
    d2 = DOCX("./example/test2.docx")
    find_and_replace(d2, "(\w+) (\w+)", r"First: \1, Last: \2")
    d2.save("./example/output2.docx")

## Other Functions
1. get_font_name and set_font_name (including chinese, arabics and others font name)

    For some reason python-docx don't support obtaining chinese, arabics and other font name of text-containing object unless going down to the _element level. I wrote get_font_name, change_font and change_font_from_change_dict to make it easier at least for me.

2. set_run_text
3. add_run_after_run
4. remove_run
5. split_run_at_string_index
6. delete_paragraph
7. isolate_para_runs_by_span
