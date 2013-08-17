function! excel_vba_echodoc#echodoc()
  if !exists("g:loaded_echodoc") 
    execute ":EchoDocEnable"
    call excel_vba_echodoc#echodoc()
	endif
endfunction

let s:doc_dict = {
  \ 'name' : 'excel_vba_echodoc',
  \ 'rank' : 100,
  \ 'filetypes' : { 'vb' : 1, 'basic' : 1 },
  \ }

function! s:doc_dict.search(cur_text)
	let ret = []
  let g:temp_cur_text = a:cur_text
  "let g:temp_result = {}
  for key in keys(s:excel_vba_text)
    "let g:temp_result[key] = matchstr(a:cur_text, key)
    let word = matchstr(a:cur_text, key)
    if word != ""
      call add(ret, { 'text': 'Usage', 'highlight': 'Identifier'})
      call add(ret, { 'text': s:excel_vba_text[word]})
    endif
  endfor
	return ret
endfunction

function! s:getfile(inputfile)"{{{
  let org = &enc
  let &enc = "utf-8"

  let ret = ""
  let ret2 = {}
  if filereadable(a:inputfile) 
    for line in readfile(a:inputfile)
      let ret = ret . line
    endfor
    let ret2 = eval(ret)
  endif

  let &enc = org
  return ret2
endfunction"}}}

function! excel_vba_echodoc#get_lib_text()"{{{
	"let path = expand("%:p:h")
	"let path = expand("<sfile>:p:h")
	let path = fnamemodify("excel_vba_echodoc.txt", ":p:h")
	let path = path . "/lib/excel_vba_echodoc.txt"
	echo path
  return path
endfunction"}}}

let s:excel_vba_text = {
  \   'MsgBox': ' MsgBox(Prompt, [Buttons], [Title], [Helpfile], [Context])',
  \   'InputBox': 
  \     ' InputBox(Prompt, [Title], [Default], [Xpos], [Ypos], [Helpfile], [Context])',
  \}

" vim:foldmethod=marker
