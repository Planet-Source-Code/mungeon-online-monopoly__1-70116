Attribute VB_Name = "mod1"
Public activeWindow As Form
Public preActiveWindow As Form
Public selectedToken As Integer

'##### Player information
Public PlayerID As Long
Public PlayerName As String
Public PlayerNumber As Integer

Public Type TokenInfo
    PID As Long
    PName As String
    tokenID As Integer
    inGame As Boolean
    currentSlot As Integer
    inParking As Boolean
    inJail As Boolean
    numTurnInJail As Integer
    cardOutOfJailA As Boolean
    cardOutOfJailB As Boolean
    cash As Currency
    ready As Boolean
    keyPlayer As Boolean
End Type
Public player(1 To 4) As TokenInfo

'##### Game Setting
Public tableID As Long
Public gameTitle As String
Public maxPlayer As Integer
Public currentRules As Integer
Public currentPlayer As Integer
Public gameStarted As Boolean
Public gameTurn As Long
Public currentChanceCard As Integer
Public chanceCard(1 To 16) As Integer
Public currentCommunityCard As Integer
Public communityCard(1 To 16) As Integer
Public housesAvailable As Integer
Public hotelsAvailable As Integer
Public rndChance(1 To 16) As Integer
Public rndCommunity(1 To 16) As Integer

Public Type GameInfo
    housesPerHotel As Integer
    totalHouses As Integer
    totalHotels As Integer
    initialCash As Long
    salary As Long
    auctionTimeDelay As Integer
    startingProperties As Integer
End Type
Public gameRules(3) As GameInfo

'##### Board Infoamtion
Public Type DeedInfo
    deedID As Integer
    number As Integer
    titleDeed As String
    color As String
    price As Currency
    rentHouse(0 To 4) As Currency
    rentWithHotel As Currency
    mortgageValue As Currency
    houseCost As Currency
    hotelCost As Currency
    currentRentType As Integer
    soundFile As String
End Type
Public deed(1 To 40) As DeedInfo

Public Type SlotInfo
    deedID As Integer
    onMortgage As Boolean
    hasOwner As Boolean
    ownerPos As Integer
    tokenSlot(1 To 4) As Boolean
    numOfHouses As Integer
    numOfHotels As Integer
End Type
Public slot(1 To 40) As SlotInfo

Public Type tokenData
    tokenID As Integer
    name As String
    file As String
    description As String
End Type
Public token(1 To 10) As tokenData

Public Sub ResetPlayerStatus()
    For i = 1 To 4
        player(i).PID = 0
        player(i).PName = ""
        player(i).tokenID = 0
        player(i).inGame = False
        player(i).currentSlot = 1
        player(i).inParking = False
        player(i).inJail = False
        player(i).numTurnInJail = 0
        player(i).cardOutOfJailA = False
        player(i).cardOutOfJailB = False
        player(i).cash = 0
        player(i).ready = False
        player(i).keyPlayer = False
    Next
End Sub

Public Sub loadDeedCardInfo()
    Dim tempStr As String
    Open "deed" For Input As #1
        For i = 1 To 40
            Input #1, tempStr
            deed(i).deedID = Split(tempStr, ";")(0)
            deed(i).number = Split(tempStr, ";")(1)
            deed(i).titleDeed = Split(tempStr, ";")(2)
            deed(i).color = Split(tempStr, ";")(3)
            deed(i).price = Split(tempStr, ";")(4)
            deed(i).rentHouse(0) = Split(tempStr, ";")(5)
            deed(i).rentHouse(1) = Split(tempStr, ";")(6)
            deed(i).rentHouse(2) = Split(tempStr, ";")(7)
            deed(i).rentHouse(3) = Split(tempStr, ";")(8)
            deed(i).rentHouse(4) = Split(tempStr, ";")(9)
            deed(i).rentWithHotel = Split(tempStr, ";")(10)
            deed(i).mortgageValue = Split(tempStr, ";")(11)
            deed(i).houseCost = Split(tempStr, ";")(12)
            deed(i).hotelCost = Split(tempStr, ";")(13)
            deed(i).currentRentType = Split(tempStr, ";")(14)
            deed(i).soundFile = Split(tempStr, ";")(15)
        Next i
    Close #1
End Sub

Public Sub rulesSetting()
    gameRules(0).housesPerHotel = 5
    gameRules(0).totalHouses = 32
    gameRules(0).totalHotels = 16
    gameRules(0).initialCash = 1500
    gameRules(0).salary = 200
    gameRules(0).auctionTimeDelay = 5
    gameRules(0).startingProperties = 0
    
    gameRules(1).housesPerHotel = 4
    gameRules(1).totalHouses = 32
    gameRules(1).totalHotels = 16
    gameRules(1).initialCash = 1500
    gameRules(1).salary = 200
    gameRules(1).auctionTimeDelay = 5
    gameRules(1).startingProperties = 2
End Sub

Public Sub loadToken()
    token(1).tokenID = 1
    token(1).name = "BATTLESHIP"
    token(1).file = "battleship.gif"
    token(1).description = "An aggressive contestant you see every game as all-out war. Inevitably your playing style will always make waves – you’re intent on building a property empire and no one’s going to stop you. But will your course to victory be an epic battle or plain sailing? We’ll see."

    token(2).tokenID = 2
    token(2).name = "CANNON"
    token(2).file = "cannon.gif"
    token(2).description = "Constantly aiming to be a big noise in the property world you see yourself as a player of the highest caliber. Your business dealings are conducted with almost military precision – you guard against the unknown and place as much emphasis on defensive strategy as you do in tactical attack. Ultimately, and befittingly, you’re far more interested in boom than bust. What better token that the cannon to get you there?"

    token(3).tokenID = 3
    token(3).name = "DOG"
    token(3).file = "dog.gif"
    token(3).description = "The tenacity and courage of the terrier is a fine metaphor for your playing style, and that’s reason enough to choose the dog. Though you’re almost certainly an animal lover, when playing Monopoly, you are anything but man’s best friend. Opponents had better be on their guard because once you’re off your leash you’ll be hard to catch, running around the board on your way to owning it all."

    token(4).tokenID = 4
    token(4).name = "HORSE"
    token(4).file = "horse.gif"
    token(4).description = "Fancying yourself as hot favorite to win any contest, you are naturally inclined to choose the horse even though you know there will be a number of hurdles on the way to victory. Any one of these could upset your chances, but your cool head allows you to stay focused on the ultimate goal and, providing you don’t let go of the reins, you should end up with your nose in front. It would take a brave opponent to bet against you."

    token(5).tokenID = 5
    token(5).name = "IRON"
    token(5).file = "iron.gif"
    token(5).description = "It’s not so much that you want to completely flatten your opponents – through of course that is what you’ll need to do, to own it all. It’s more that you’re habitually neat and tidy in all your property and financial dealings, and that you like things to go smoothly. For you, the only thing that rightfully belongs on a board is an iron."

    token(6).tokenID = 6
    token(6).name = "RACE CAR"
    token(6).file = "car.gif"
    token(6).description = "An extremely confident sort, you only know one way to play Monopoly – fast! You drive hard deals in all your property negotiations and, try as they might, your fellow players struggle to keep up. You just can’t wait to build up your property stable and you’ll spend what you have to, to own it all. The only question is, do your have the skills to stay on track?"

    token(7).tokenID = 7
    token(7).name = "SHOE"
    token(7).file = "shoe.gif"
    token(7).description = "You’ve trodden your way around the classic Monopoly board so many times you could find your way in the dark. Still, there’s no substitute for experience and the wisdom that comes from hours of game play makes you a canny contestant indeed. Your sometimes scruffy appearance masks a player whose property dealing is methodical and focused. Each step you take is purposeful progress on the road to owning it all."
    
    
    token(8).tokenID = 8
    token(8).name = "THIMBLE"
    token(8).file = "thimble.gif"
    token(8).description = "Cautions by nature, you prefer not to risk spending your money too quickly. You’re only too aware that a slip up in your money management could lead to painful ruination. Then again, you’re aware of the need to speculate to accumulate. Ultimately, buying well within your means, your prudence is your greatest playing strength."
    
    
    token(9).tokenID = 9
    token(9).name = "TOP HAT"
    token(9).file = "hat.gif"
    token(9).description = "There’s no mistaking your aspirations. Like Mr. Monopoly himself, your ambition is simple – to own it all. Your taste for the finer things in life can lead to a preoccupation with the more valuable properties on the board and you are unlikely to be satisfied with anything you purchase until it is fully developed. You resent paying your fines and taxes as much as the next player, but are magnanimous in handing over rent and birthday presents."
    
    
    token(10).tokenID = 10
    token(10).name = "WHEELBARROW"
    token(10).file = "wheelbarrow.gif"
    token(10).description = "Perhaps in anticipation of winning barrow-loads of money off your fellow players you choose the wheelbarrow. With a firm grip on the handles you’re unlikely to let commercial opportunities pass you by. Maneuverable, spacious, and designed to cope with bumps along the way, the course you steer is inexorably towards owning it all."
End Sub

