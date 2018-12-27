option explicit

sub main() ' {

    dim sqlStmt as string

    sqlStmt = "create table tab_xyz (" & vbcrlf & _
              "  col_1 integer,"       & vbcrlf & _
              "  col_2 integer"        & vbcrlf & _
              ")"

    application.currentDB().execute(sqlStmt)

end sub ' }
