dbMemo "SQL" ="TRANSFORM First(\"X\") AS Wyr1\015\012SELECT projekty.nazwaProjektu, projekty.et"
    "apWLejku, projekty.ID AS IDProjektu\015\012FROM lejki LEFT JOIN projekty ON lejk"
    "i.ID = projekty.etapWLejku\015\012WHERE (((projekty.status)=1 Or (projekty.statu"
    "s) Is Null))\015\012GROUP BY projekty.nazwaProjektu, projekty.etapWLejku, projek"
    "ty.ID, projekty.status\015\012ORDER BY projekty.nazwaProjektu, lejki.ID\015\012P"
    "IVOT lejki.ID;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbByte "PublishToWeb" ="1"
        dbText "Name" ="projekty.nazwaProjektu"
        dbLong "AggregateType" ="-1"
    End
        dbText "Name" ="Wyr1"
    End
        dbText "Name" ="IDProjektu"
    End
        dbText "Name" ="1"
    End
        dbText "Name" ="2"
    End
        dbText "Name" ="3"
    End
        dbText "Name" ="4"
    End
        dbText "Name" ="5"
    End
        dbText "Name" ="6"
    End
End
