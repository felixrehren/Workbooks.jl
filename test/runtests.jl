using Test, Workbooks

###  W O R K B O O K S
wb = Workbook()
S = wb["Sheet1"]

@testset "references       " begin
    @test Workbooks.LocalPosition(1,1) == Workbooks.LocalPosition("A1")
    @test Workbooks.LocalPosition(1,1) == Workbooks.LocalPosition("A",1)
    @test Workbooks.LocalPosition(1,1) == Workbooks.LocalRef("A1")

    @test Workbooks.GlobalPosition("Sheet1!A1") == Workbooks.GlobalPosition("Sheet1","A1")
end

@testset "basic formulas   " begin
    # test constant assignment
    S["A1"] = "3"
    @test S["A1"].value == 3
    # testing constant formulas
    S["A2"] = "=5"
    @test S["A2"].value == 5
    # the next value will be rewritten
    S["B1"] = "7"
    # test a formula
    S["B2"] = "=1+A1*(B1-A2)" # now = 1 + 3*(7-5) = 7
    @test S["B2"].value == 7
    # testing formulas without inputs
    S["D1"] = "=6*6"
    @test S["D1"].value == 36
    # testing formulas with missing inpuits
    S["C3"] = "=B2/B3"
    @test ismissing(S["C3"].value)
    # supplying the missing input
    S["B3"] = "2"
    @test S["C3"].value == 7/2    # C3 should now be 7/2 = 3.5
    # refreshing an existing input
    S["B1"] = "11"
    @test S["B2"].value == 19     # B2 = 1 + 3*(11-5) = 19
    @test S["C3"].value == 19/2   # C3 = 19/2 = 9.5
    # test repeated references
    S["A3"] = "=1/A1 + 2/A2 + 3/A1"
    @test S["A3"].value == 1/3 + 2/5 + 3/3
    # test simplest reference
    S["C1"] = "=A3"
    @test S["C1"].value == S["A3"].value
end

@testset "evaluating ranges" begin
    # reference to a range
    S["D2"] = "=A1:A3"
    @test all(S["D2"].value .== Real[S["A1"].value, S["A2"].value, S["A3"].value])
    # summing the range
    S["D3"] = "=sum(D2)"
    @test S["D3"].value == sum(getfield.(S["A1:A3"],:value))
    @test S["D3"].value == (9 + 2/5 + 1/3)
    # refreshing with a range reference
    S["A3"] = "7"
    @test S["D3"].value == 3 + 5 + 7 
end

@testset "cross-sheet refs " begin
    S2 = wb["Sheet2"]
    S2["A1"] = "1"
    S2["A2"] = "=Sheet1!A2"
    @test S2["A2"].value == S["A2"].value
end

@testset "file system      " begin
    Workbooks.write(wb)
    wb2 = Workbooks.read()
    @test all(skipmissing(Workbooks.array(S) .== Workbooks.array(wb2["Sheet1"])))
end