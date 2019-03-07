option explicit

function nvl(v as variant) as string ' {

    if v is nothing then
       nvl = ""
       exit function
    end if

    nvl = v

end function ' }

sub langs() ' {

    dim fileName as string
    fileName = environ$("temp") & "\languages.txt"

    dim f as integer
    f = freeFile()

    open fileName for output as #f

    dim main_lang_id as long
    dim sub_lang_id  as long

    dim lang as language
    for each lang in languages ' {

        sub_lang_id  = lang.id  mod  1024
        main_lang_id = lang.id  \    1024

        print# f, sub_lang_id    & chr(9) & _
                  main_lang_id   & chr(9) & _
                  lang.id        & chr(9) & _
                  lang.nameLocal
                ' chr(9) & nvl(lang.activeGrammarDictionary    ) &
                ' chr(9) & nvl(lang.activeHyphenationDictionary) &
                ' chr(9) & nvl(lang.activeSpellingDictionary   ) & _
                ' chr(9) & nvl(lang.activeThesaurusDictionary  )
    next lang ' } 

    close# f

end sub ' }
