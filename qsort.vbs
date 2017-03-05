option explicit
'on error resume next

sub initArr(byref a(), n)
    randomize timer
    dim i
    for i = 0 to n - 1
        a(i) = int(100 * rnd + 1)
    next
end sub

sub printArr(a())
    if ubound(a) <= 25 then
        dim num
        for each num in a
            wscript.stdout.write num & " "
        next
        wscript.echo ""
    end if
end sub

sub swap(byref a(), i, j)
    dim t
    'wscript.echo "before swap ij is " & i & " " & j
    t = a(i)
    a(i) = a(j)
    a(j) = t
end sub

sub median3(byref a, lo, hi)
    'goal: a(m) <= a(lo) <= a(hi) so median is in a(lo)
    dim md
    md = lo + (hi - lo) \ 2
    if a(lo) > a(hi) then
        swap a, lo, hi
    end if
    if a(md) > a(hi) then
        swap a, md, hi
    end if
    if a(md) > a(lo) then
        swap a, lo, md
    end if
    'wscript.echo "Median3 " & a(lo) & " " & a(md) & " " & a(hi)
end sub

function partition1(byref a(), lo, hi)
    dim i, j, p
    i = lo
    j = hi + 1
    'swap a, lo, lo + (hi - lo) \ 2 'pivot is middle moved to position lo
    median3 a, lo, hi
    p = a(lo)
    'wscript.echo "p is " & p
    'printarr a
    do
        do
            i = i + 1
            if i = hi then
                exit do
            end if
        loop while a(i) <= p
        do
            j = j - 1
        loop while a(j) > p 'will always stop at pivot p
        if i >= j then
            exit do
        end if
        swap a, i, j
        'printArr a
    loop while true
    swap a, lo, j 'move pivot to correct location at j
    'wscript.echo "final ij is " & i & " " & j
    'printArr a
    partition1 = j
end function

function partition2(byref a(), lo, hi)
    dim i, j, p
    j = lo
    'swap a, lo, lo + (hi - lo) \ 2 'pivot is middle moved to position lo
    median3 a, lo, hi
    p = a(lo)
    'wscript.echo "p is " & p
    'printarr a
    for i = lo + 1 to hi
        if a(i) <= p then
            j = j + 1
            swap a, i, j
            'printArr a
        end if
    next
    swap a, lo, j 'move pivot to correct location at j
    'wscript.echo "final ij is " & i & " " & j
    'printArr a
    partition2 = j
end function

function partition3(byref a(), lo, hi)
    dim i, j, p
    i = lo - 1
    j = hi + 1
    median3 a, lo, hi
    'p = a(lo + (hi - lo) \ 2)
    p = a(lo)
    'wscript.echo "p is " & p & " " & lo & " " & hi
    'printarr a
    do
        do
            i = i + 1
            if i = hi then
                exit do
            end if
        loop while a(i) < p
        do
            j = j - 1
        loop while a(j) > p 'will always stop at pivot p
        if i >= j then
            exit do
        end if
        swap a, i, j
        'printArr a
    loop while true
    'wscript.echo "final ij is " & i & " " & j
    'printArr a
    partition3 = j
end function

sub myqsort(byref a(), lo, hi)
    dim i, j, p, t                          'median 3 w pivot at lo
    if lo >= hi then: exit sub :end if
    if (hi - lo) < 10 then 'do insert sort
        for i = lo+1 to hi
            t = a(i)
            j = i
            do while j > 0 'no short circuit logic in vbscript :-(
                if a(j-1) > t then
                    a(j) = a(j-1)
                    j = j-1
                else
                    exit do
                end if
            loop
            a(j) = t
        next
        'printarr a
        exit sub
    end if
    'printarr a
    'wscript.echo "lo " & lo & " hi " & hi
    p = lo + (hi - lo) \ 2
    t = a(lo+1) : a(lo+1) = a(p) : a(p) = t 'swap middle and lo+1
    if a(lo) > a(hi) then
        t = a(lo) : a(lo) = a(hi) : a(hi) = t
    end if
    if a(lo+1) > a(lo) then
        t = a(lo+1) : a(lo+1) = a(lo) : a(lo) = t
    end if
    if a(lo) > a(hi) then
        t = a(lo) : a(lo) = a(hi) : a(hi) = t 'lo+1 <= lo <= hi
    end if
    i = lo+1 : j = hi : p = a(lo)           'initialize loop variables
    'wscript.echo "p is " & p
    'printarr a
    do
        do: i = i+1 :loop while a(i) < p    'will always stop at hi
        do: j = j-1 :loop while a(j) > p    'will always stop at lo+1
        if i >= j then: exit do :end if
        t = a(i) : a(i) = a(j) : a(j) = t   'swap i and j
        'wscript.echo "swap ij is " & i & " " & j
        'printArr a
    loop while true
    a(lo) = a(j) : a(j) = p     'move pivot to it's slot at j
    'wscript.echo "final ij is " & i & " " & j
    'printArr a
    myqsort a, lo, j-1
    myqsort a, j+1, hi
end sub

sub qsort(byref a(), lo, hi)
    if lo < hi then
        dim q
        q = partition2(a, lo, hi) 'partition1 or 2 is compatable
        qsort a, lo, q-1
        qsort a, q+1, hi
    end if
end sub

sub qsort2(byref a(), lo, hi)
    if lo < hi then
        dim q
        q = partition3(a, lo, hi) 'only partition3 is compatable
        qsort2 a, lo, q
        qsort2 a, q + 1, hi
    end if
end sub

sub insertsort(byref a(), n)
    dim i, j, t
    for i = 1 to n - 1
        t = a(i)
        j = i
        do while j > 0
            if a(j - 1) > t then
                a(j) = a(j - 1)
                j = j - 1
            else
                exit do
            end if
        loop
        a(j) = t
        'wscript.echo "final ij is " & i & " " & j
        'printArr a
    next
end sub

sub shellsort(byref a(), n)
    dim i, j, t, gap, gaps
    gaps = array(797161, 265720, 88573, 29524, 9841, 3280, 1092, 364, 121, 40, 13, 4, 1)
    for each gap in gaps
        'wscript.echo "gap is " & gap
        for i = gap to n-1
            t = a(i)
            j = i
            do while (j-gap) >= 0
                if a(j-gap) > t then
                    a(j) = a(j - gap)
                    j = j - gap
                else
                    exit do
                end if
            loop
            a(j) = t
            'wscript.echo "final ij is " & i & " " & j
            'printArr a
        next
    next
end sub


dim n, a(), timstart, timend
'a = array(67, 88, 92, 26, 50, 93, 40, 8, 48, 96)
'a=array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10)
n = 100000
redim preserve a(n)
initArr a, n
printArr a
timstart = timer
'qsort a, 0, n - 1
'qsort2 a, 0, n - 1
'insertsort a, n
'shellsort a, n
myqsort a, 0, n-1
timend = timer
printArr a
wscript.echo timend - timstart
