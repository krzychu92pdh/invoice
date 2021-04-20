def generator(liczba):

    liczba = str(liczba).split(".")
    liczba1 = list(liczba[0])
    liczba1.reverse()

    slownie = []

    jednosci = ['zero', 'jeden', 'dwa', 'trzy', 'cztery', 'pięć', 'sześć', 'siedem', 'osiem', 'dziewięć']
    dziesiatki = ['', 'dziesięć', 'dwadzieścia', 'trzydzieści', 'czterdzieści', 'pięćdziesiąt', 'sześćdziesiąt', 'siedemdziesiąt', 'osiemdziesiąt', 'dziewięćdziesiąt']
    dziesiatki1 = ['', 'jedenaście', 'dwanaście', 'trzynaście', 'czternaście', 'piętnaście', 'szesnaście', 'siedemnaście', 'osiemnascie', 'dziewiętnaście']
    setki = ['', 'sto', 'dwieście', 'trzysta', 'czterysta', 'pięćset', 'sześćset', 'siedemset', 'osiemset', 'dziewięćset']
    tysiace = ['', 'tysiąc', 'dwa tysiące', 'trzy tysiące', 'cztery tysiące', 'pięć tysięcy', 'sześć tysięcy', 'siedem tysięcy', 'osiem tysięcy', 'dziewięć tysięcy']


    ile = len(liczba1)

    pierwszy = liczba1[0]
    pierwsza_liczba = int(pierwszy)
    if ile==2 or ile == 3 or ile ==4:
        drugi = liczba1[1]
        druga_liczba = int(drugi)
    if ile == 3 or ile ==4:
        trzeci = liczba1[2]
        trzecia_liczba = int(trzeci)
    if ile == 4:
        czwarty = liczba1[3]
        czwarta_liczba = int(czwarty)


    if ile==2 or ile == 3 or ile ==4:
        if druga_liczba != 1 and pierwsza_liczba != 0:
            slownie.insert(0,jednosci[pierwsza_liczba])
        if druga_liczba != 0:
            if druga_liczba == 1:
                slownie.insert(1, dziesiatki1[pierwsza_liczba])
            else:
                slownie.insert(1, dziesiatki[druga_liczba])

    if ile == 3 or ile ==4:        
        if trzecia_liczba != 0:
            slownie.insert(2,setki[trzecia_liczba])

    
    if ile ==4:
        if czwarta_liczba != 0:
            slownie.insert(3,tysiace[czwarta_liczba])

    x = int(liczba1[0])
    if x == 2 or x==3 or x ==4:
        slownie.insert(0, "złote")
    else:
        slownie.insert(0, "złotych")

    slownie.reverse()

    s=" "
    s = s.join(slownie)
    return s
