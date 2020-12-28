using Test, Workbooks

@testset "Sylvaticus' example" begin
    out = plusTwo(3)
    @test out == 5
end

###  W O R K B O O K S
wb = Workbook()
S = wb["Sheet1"]

@testset "basic formulas" begin
    # test constant assignment
    set!(wb,"Sheet1!A1","3")
    @test S["A1"].value == 3
    # testing constant formulas
    set!(wb,"Sheet1!A2","=5") 
    @test S["A2"].value == 5
    # value will be rewritten
    set!(wb,"Sheet1!B1","7") 
    # test a formula
    set!(wb,"Sheet1!B2","=1+A1*(B1-A2)") # now = 1 + 3*(7-5) = 7
    @test S["B2"].value == 7
    # testing formulas without inputs
    set!(wb,"Sheet1!D1","=6*6")
    @test S["D1"].value == 36
    # testing formulas with missing inpuits
    set!(wb,"Sheet1!C3","=B2/B3")
    @test ismissing(S["C3"].value)
    # supplying the missing input
    set!(wb,"Sheet1!B3","2")
    @test S["C3"].value == 7/2    # C3 should now be 7/2 = 3.5
    # refreshing an existing input
    set!(wb,"Sheet1!B1","11") 
    @test S["B2"].value == 19     # B2 = 1 + 3*(11-5) = 19
    @test S["C3"].value == 19/2   # C3 = 19/2 = 9.5
    # test repeated references
    set!(wb,"Sheet1!A3","=1/A1 + 2/A2 + 3/A1")
    @test S["A3"].value == 1/3 + 2/5 + 3/3
end