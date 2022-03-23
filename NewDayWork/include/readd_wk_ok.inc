<%

'新增資料之SQL指令字串
strSQL_add="Insert into "&tb_name&" ("
strSQL_add=strSQL_add & "undo_date1,doing_date1,wk_class,wk_group,"
strSQL_add=strSQL_add & "wk_item,wk_content,wk_order,"
strSQL_add=strSQL_add & "wk_doer,wk_undoer) values ('"

strSQL_add=strSQL_add &Request("undo_date1")&"','"
strSQL_add=strSQL_add &Request("doing_date1")&"','"
strSQL_add=strSQL_add &Request("wk_class")&"','"
strSQL_add=strSQL_add &Request("wk_group")&"','"

strSQL_add=strSQL_add &Request("wk_item")&"','"
strSQL_add=strSQL_add &Request("wk_content")&"','"
strSQL_add=strSQL_add &Request("wk_order")&"','"

strSQL_add=strSQL_add &Request("all_worker")&"','"
strSQL_add=strSQL_add &Request("all_worker")&"')"

'執行SQL指令
conDB.Execute strSQL_add
%>
