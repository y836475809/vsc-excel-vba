Open fp For Output Access Read Write Lock Read Write As #1

Open "TESTFILE" For Output As #1
Write #1, "write test", 234
Write #1,
Write #1, value ; "write value"

Print #1, "print test"
Print #1,
Print #1, "print 1"; Tab ; "print 2" 
Print #1, "Hello" ; " " ; "World" 
Print #1, Spc(5) ; "5spaces"
Print #1, Tab(10, 1, Spc(3)) ; "10tab"

Write fn, text

