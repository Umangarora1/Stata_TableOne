
program define export_table
    version 14
    syntax anything, BY(varname) [MEDIANONLY]

    quietly putexcel set "Table Export Group.xlsx", modify

    tempvar __by
    local bylab ""
    capture confirm numeric variable `by'
    if _rc {
        encode `by', gen(`__by')
        local bylab : value label `__by'
    }
    else {
        gen double `__by' = `by'
        local bylab : value label `by'
    }

    quietly levelsof `__by' if !missing(`__by'), local(BYL)
    local BYcount : word count `BYL'
    if `BYcount' != 2 {
        di as error "by() must have exactly two levels (found `BYcount')."
        exit 198
    }
    local BY0 : word 1 of `BYL'
    local BY1 : word 2 of `BYL'

    local head0 = "`by' `=string(`BY0')'"
    local head1 = "`by' `=string(`BY1')'"
    if "`bylab'" != "" {
        capture local g0 : label (`bylab') `BY0'
        capture local g1 : label (`bylab') `BY1'
        if "`g0'" != "" local head0 "`g0'"
        if "`g1'" != "" local head1 "`g1'"
    }

    quietly putexcel A1=("Variable") B1=("`head0'") C1=("N `head0'") ///
        D1=("`head1'") E1=("N `head1'") F1=("P value") G1=("Stat test") ///
        H1=("Total") I1=("N total")

    *---------------- parse "anything" to find i.-flagged variables -------------
    local raw `"`anything'"'
    local ALLV ""
    local CATFORCE ""
    tokenize `"`raw'"'
    while "`1'" != "" {
        local tok "`1'"
        macro shift
        if substr("`tok'",1,2)=="i." {
            local vn = substr("`tok'",3,.)
            capture unab vn2 : `vn'
            if !_rc {
                local CATFORCE "`CATFORCE' `vn2'"
                local ALLV "`ALLV' `vn2'"
            }
        }
        else {
            capture unab vn2 : `tok'
            if !_rc local ALLV "`ALLV' `vn2'"
        }
    }
    local VARS : list uniq ALLV
    local row = 2

    foreach v of local VARS {

        quietly count if !missing(`v')
        if (r(N)==0) continue

        *---------------- decide categorical vs continuous -----------------------
        local iscat 0
        local isforced : list v in CATFORCE
        if `isforced' {
            local iscat 1
        }
        else {
            capture confirm numeric variable `v'
            if _rc {
                quietly levelsof `v' if !missing(`v'), local(Ls)
                local Lc : word count `Ls'
                if `Lc'==2 local iscat 1
                else local iscat 1   /* string with >2 levels -> categorical for safety */
            }
            else {
                quietly levelsof `v' if !missing(`v'), local(Ln)
                local Lc : word count `Ln'
                if `Lc'==2 local iscat 1
                else local iscat 0   /* >2 numeric levels and not forced -> continuous */
            }
        }

        if `iscat'==1 {

            quietly count if `__by'==`BY0' & !missing(`v')
            local N0 = r(N)
            quietly count if `__by'==`BY1' & !missing(`v')
            local N1 = r(N)
            quietly count if !missing(`v')
            local NT = r(N)

            * p-value (Fisher only if 2x2); use numeric temp for safety
            tempvar __catnum
            capture confirm string variable `v'
            if !_rc {
                encode `v', gen(`__catnum')
            }
            else {
                gen double `__catnum' = `v'
            }

            tempname PCHI
            scalar `PCHI' = .
            local testlab "chi2"
            quietly levelsof `__catnum' if !missing(`__catnum'), local(L_all)
            local Lcount_all : word count `L_all'
            if `Lcount_all'==2 {
                quietly tab `__catnum' `__by' if !missing(`__catnum') & ///
                    inlist(`__by',`BY0',`BY1'), exact
                scalar `PCHI' = r(p)
                local testlab "Fisher exact"
            }
            if missing(`PCHI') {
                quietly tab `__catnum' `__by' if !missing(`__catnum') & ///
                    inlist(`__by',`BY0',`BY1'), chi
                scalar `PCHI' = r(p)
                local testlab "chi2"
            }
            local ptext ""
            if !missing(`PCHI') {
                if `PCHI' < .001 local ptext "<0.001"
                else if `PCHI' < .01 local ptext = string(`PCHI',"%9.3f")
                else local ptext = string(`PCHI',"%9.2f")
            }

            quietly levelsof `v' if !missing(`v'), local(L)

            * collapse only 0/1 numerics to show the "=1" row
            local collapse01 0
            capture confirm numeric variable `v'
            if !_rc {
                quietly levelsof `v' if !missing(`v'), local(Ln)
                local Lc : word count `Ln'
                if `Lc'==2 {
                    local a : word 1 of `Ln'
                    local b : word 2 of `Ln'
                    capture confirm number `a'
                    if !_rc {
                        if inlist(`a',0,1) & inlist(`b',0,1) local collapse01 1
                    }
                }
            }
            if `collapse01' local L 1

            local firstrow 1
            foreach lvl of local L {

                capture confirm numeric variable `v'
                if _rc {
                    quietly count if `__by'==`BY0' & `v'=="`lvl'"
                    local n0 = r(N)
                    quietly count if `__by'==`BY1' & `v'=="`lvl'"
                    local n1 = r(N)
                    quietly count if `v'=="`lvl'"
                    local nT = r(N)
                    local lvlname "`lvl'"
                }
                else {
                    quietly count if `__by'==`BY0' & `v'==`lvl'
                    local n0 = r(N)
                    quietly count if `__by'==`BY1' & `v'==`lvl'
                    local n1 = r(N)
                    quietly count if `v'==`lvl'
                    local nT = r(N)
                    local lvlname "`lvl'"
                    local vallab : value label `v'
                    if "`vallab'"!="" {
                        capture local labtxt : label (`vallab') `lvl'
                        if "`labtxt'"!="" local lvlname "`labtxt'"
                    }
                }

                local p0 = .
                local p1 = .
                local pT = .
                if `N0' > 0 local p0 = 100*`n0'/`N0'
                if `N1' > 0 local p1 = 100*`n1'/`N1'
                if `NT' > 0 local pT = 100*`nT'/`NT'

                local vlabel : variable label `v'
                if "`vlabel'"=="" local vlabel "`v'"
                local rowname "`vlabel'"
                if !`collapse01' local rowname "`vlabel': `lvlname'"

                local Btxt = string(`n0',"%9.0f") + " (" + string(`p0',"%9.2f") + "%)"
                local Dtxt = string(`n1',"%9.0f") + " (" + string(`p1',"%9.2f") + "%)"
                local Htxt = string(`nT',"%9.0f") + " (" + string(`pT',"%9.2f") + "%)"

                if `firstrow' {
                    quietly putexcel A`row'=("`rowname'") B`row'=("`Btxt'") ///
                        C`row'=(`N0') D`row'=("`Dtxt'") E`row'=(`N1') ///
                        F`row'=("`ptext'") G`row'=("`testlab'") ///
                        H`row'=("`Htxt'") I`row'=(`NT')
                }
                else {
                    quietly putexcel A`row'=("`rowname'") B`row'=("`Btxt'") ///
                        C`row'=(`N0') D`row'=("`Dtxt'") E`row'=(`N1') ///
                        H`row'=("`Htxt'") I`row'=(`NT')
                }

                local row = `row' + 1
                local firstrow = 0
            }
        }
        else {

            *--------------- continuous -------------------
            if "`medianonly'"!="" {
                quietly ranksum `v', by(`__by')
                scalar pv = 2*normprob(-abs(r(z)))
                local ptext ""
                if !missing(pv) {
                    if pv < .001 local ptext "<0.001"
                    else if pv < .01 local ptext = string(pv,"%9.3f")
                    else local ptext = string(pv,"%9.2f")
                }
                local testlab "ranksum"

                quietly tabstat `v' if `__by'==`BY0', s(n p50 p25 p75) save
                matrix M0 = r(StatTotal)
                local n0 = string(M0[1,1],"%9.0f")
                local md0 = string(M0[2,1],"%9.2f")
                local q10 = string(M0[3,1],"%9.2f")
                local q30 = string(M0[4,1],"%9.2f")

                quietly tabstat `v' if `__by'==`BY1', s(n p50 p25 p75) save
                matrix M1 = r(StatTotal)
                local n1 = string(M1[1,1],"%9.0f")
                local md1 = string(M1[2,1],"%9.2f")
                local q11 = string(M1[3,1],"%9.2f")
                local q31 = string(M1[4,1],"%9.2f")

                quietly tabstat `v', s(n p50 p25 p75) save
                matrix MT = r(StatTotal)
                local nT = string(MT[1,1],"%9.0f")
                local mdT = string(MT[2,1],"%9.2f")
                local q1T = string(MT[3,1],"%9.2f")
                local q3T = string(MT[4,1],"%9.2f")

                local vlabel : variable label `v'
                if "`vlabel'"=="" local vlabel "`v'"

                local Btxt = "`md0' (`q10' - `q30')"
                local Dtxt = "`md1' (`q11' - `q31')"
                local Htxt = "`mdT' (`q1T' - `q3T')"

                quietly putexcel A`row'=("`vlabel'") B`row'=("`Btxt'") ///
                    C`row'=("`n0'") D`row'=("`Dtxt'") E`row'=("`n1'") ///
                    F`row'=("`ptext'") G`row'=("`testlab'") ///
                    H`row'=("`Htxt'") I`row'=("`nT'")
                local row = `row' + 1
            }
            else {
                local normal0 = .
                local normal1 = .
                quietly swilk `v' if `__by'==`BY0'
                if !_rc local normal0 = (r(p)>0.05)
                quietly swilk `v' if `__by'==`BY1'
                if !_rc local normal1 = (r(p)>0.05)

                if (`normal0'==1 & `normal1'==1) {

                    quietly ttest `v', by(`__by')
                    local ptext ""
                    if !missing(r(p)) {
                        if r(p) < .001 local ptext "<0.001"
                        else if r(p) < .01 local ptext = string(r(p),"%9.3f")
                        else local ptext = string(r(p),"%9.2f")
                    }
                    local testlab "ttest"

                    quietly tabstat `v' if `__by'==`BY0', s(n mean sd) save
                    matrix T0 = r(StatTotal)
                    local n0  = string(T0[1,1],"%9.0f")
                    local m0  = string(T0[2,1],"%9.2f")
                    local sd0 = string(T0[3,1],"%9.2f")

                    quietly tabstat `v' if `__by'==`BY1', s(n mean sd) save
                    matrix T1 = r(StatTotal)
                    local n1  = string(T1[1,1],"%9.0f")
                    local m1  = string(T1[2,1],"%9.2f")
                    local sd1 = string(T1[3,1],"%9.2f")

                    quietly tabstat `v', s(n mean sd) save
                    matrix TT = r(StatTotal)
                    local nT  = string(TT[1,1],"%9.0f")
                    local mT  = string(TT[2,1],"%9.2f")
                    local sdT = string(TT[3,1],"%9.2f")

                    local vlabel : variable label `v'
                    if "`vlabel'"=="" local vlabel "`v'"

                    local Btxt = "`m0' (±`sd0')"
                    local Dtxt = "`m1' (±`sd1')"
                    local Htxt = "`mT' (±`sdT')"

                    quietly putexcel A`row'=("`vlabel'") B`row'=("`Btxt'") ///
                        C`row'=("`n0'") D`row'=("`Dtxt'") E`row'=("`n1'") ///
                        F`row'=("`ptext'") G`row'=("`testlab'") ///
                        H`row'=("`Htxt'") I`row'=("`nT'")
                    local row = `row' + 1
                }
                else {

                    quietly ranksum `v', by(`__by')
                    scalar pv = 2*normprob(-abs(r(z)))
                    local ptext ""
                    if !missing(pv) {
                        if pv < .001 local ptext "<0.001"
                        else if pv < .01 local ptext = string(pv,"%9.3f")
                        else local ptext = string(pv,"%9.2f")
                    }
                    local testlab "ranksum"

                    quietly tabstat `v' if `__by'==`BY0', s(n p50 p25 p75) save
                    matrix M0 = r(StatTotal)
                    local n0 = string(M0[1,1],"%9.0f")
                    local md0 = string(M0[2,1],"%9.2f")
                    local q10 = string(M0[3,1],"%9.2f")
                    local q30 = string(M0[4,1],"%9.2f")

                    quietly tabstat `v' if `__by'==`BY1', s(n p50 p25 p75) save
                    matrix M1 = r(StatTotal)
                    local n1 = string(M1[1,1],"%9.0f")
                    local md1 = string(M1[2,1],"%9.2f")
                    local q11 = string(M1[3,1],"%9.2f")
                    local q31 = string(M1[4,1],"%9.2f")

                    quietly tabstat `v', s(n p50 p25 p75) save
                    matrix MT = r(StatTotal)
                    local nT = string(MT[1,1],"%9.0f")
                    local mdT = string(MT[2,1],"%9.2f")
                    local q1T = string(MT[3,1],"%9.2f")
                    local q3T = string(MT[4,1],"%9.2f")

                    local vlabel : variable label `v'
                    if "`vlabel'"=="" local vlabel "`v'"

                    local Btxt = "`md0' (`q10' - `q30')"
                    local Dtxt = "`md1' (`q11' - `q31')"
                    local Htxt = "`mdT' (`q1T' - `q3T')"

                    quietly putexcel A`row'=("`vlabel'") B`row'=("`Btxt'") ///
                        C`row'=("`n0'") D`row'=("`Dtxt'") E`row'=("`n1'") ///
                        F`row'=("`ptext'") G`row'=("`testlab'") ///
                        H`row'=("`Htxt'") I`row'=("`nT'")
                    local row = `row' + 1
                }
            }
        }
    }
end
