option explicits

sub main()

    dim conn_ado as ADODB.connection
    dim conn_dao as DAO.connection

 '
 '  currentProject.connection returns an ADO
 '  connection object, not a DAO connection object.
 '  So, the following assignment works:
    set conn_ado = currentProject.connection

 '  while the following assignment would cause
 '  a "Type Mismatch" Error:
 '  set conn_dao = CurrentProject.Connection

end sub
