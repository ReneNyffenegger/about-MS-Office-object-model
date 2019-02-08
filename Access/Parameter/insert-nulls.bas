option explicit

sub main() ' {

    if not isNull(dLookup("Name", "MSysObjects", "Name='tq84_null_test'")) then
       doCmd.close acTable, "tq84_null_test", acSaveNo
       execSQL "drop table tq84_null_test"
    end if

    execSQL "create table tq84_null_test(num number, txt varchar(20), dt date)"

    insertLiterally
    insertWithParameters

    doCmd.openTable "tq84_null_test"

end sub ' }

sub insertLiterally() ' {

    execSQL "insert into tq84_null_test values (   1, 'one'  , now()                )"
    execSQL "insert into tq84_null_test values (   2,  null  , #02/22/1922 02:22:20#)"
    execSQL "insert into tq84_null_test values (null, 'three', #03/30/1933 03:33:30#)"
    execSQL "insert into tq84_null_test values (   4, 'four' , null                 )"

end sub ' }

sub insertWithParameters() ' {

    dim stmt as dao.queryDef
    set stmt = currentDB.createQueryDef("", _
      "parameters"                        & _
      "  num      number  ,"              & _
      "  txt      varchar(20),"           & _
      "  dt       date    ;"              & _
      "insert into tq84_null_test (num, txt, dt) values([num], [txt], [dt])")


      dim parNum as dao.parameter
      dim parTxt as dao.parameter
      dim parDt  as dao.parameter

      set parNum = stmt.parameters("num")
      set parTxt = stmt.parameters("txt")
      set parDt  = stmt.parameters("dt" )


      parNum =     5
      parTxt = "five"
      parDt  = #05/05/1955 05:55:50#
      stmt.execute

      parNum =    6
      parTxt = nothing               ' Use nothing, not vbNull!
      parDt  = #06/06/1966 06:06:06#
      stmt.execute

      parNum = nothing               ' Use nothing, not vbNull!
      parTxt = "seven"
      parDt  = #07/07/1977 07:07:07#
      stmt.execute

      parNum =     8
      parTxt = "eight"
      parDt  =  nothing              ' Use nothing, not vbNull!
      stmt.execute

end sub ' }

sub execSQL(stmt as string) ' {

    currentProject.connection.execute(stmt)

end sub ' }
