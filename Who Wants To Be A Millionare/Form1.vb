Public Class Form1
    'Abinash Singh
    'Who wants To Be A Millionare? 
    'Dec 7 2017

    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    ' Note: Changed Do Lopp While in Load questions to loop untill and equal sign comparison to not equal comparison program dose not freeze when ney game is clicked.
    ' also button changes color depending on if answer was right (lime) and if it was wrong (red)
    ' image used https://en.wikipedia.org/wiki/Who_Wants_to_Be_a_Millionaire%3F
    'Dont need the life line sub procedure due to only one parameter becouse it would be more efficent to proccess that in the click procedure in the button
    'Created a new sub proccedure to reset the game every time some one walks away or loses and if they run out of time (ResetGame)
    'i made a new array to store all of the labels on the form also made a new array to store all the buttons on the form (btnAllButtonsOnForm).
    'Added a start game button instead of having a blank screen at start it disapears when clicked.
    'added progeress bars (prgTimeLeft, and prgTimeRight) to illustrate time left they are located above the option buttons and under the lblquestion label
    'mutiply the length of time(60 secs) by 1.6666667 
    'Made new variable strUserResponse to register the users response in a yes no message box
    'made a btnAllButtonsOnForm arrasy to store all buttons on form for reset game sub procedure
    'i have gotten most of these questions from this website https://chartcons.com/100-hard-trivia-questions-answers/
    'got help for the list box control at http://www.vb-helper.com/howto_net_set_listbox_selections.html
    'added epilepsy when you win the game the 4 option buttons change color with a timer set to an interval of 100
    'Removed constant intMAXQUESTIONSPERGAME becouse the for loop it was beeing used in didnt work
    'Added intProgress bar divisor variable to be able to calculate value by which the progress bar will be multiplyed by to decrease the value
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

    Structure QuestionsDataBase
        Dim strQuestion As String
        Dim strCorrectAnswer As String
        Dim strHintPicture As String
        Dim strWrongAnswerOne As String
        Dim strWrongAnswerTwo As String
        Dim strWrongAnswerThree As String
    End Structure

    Structure QuestionsInUse
        Dim strQuestion As String
        Dim strAnswer As String
        Dim strHintPicture As String
        Dim strBadAnswerOne As String
        Dim strBadAnswerTwo As String
        Dim strBadAnswerThree As String
    End Structure

    'Const intMAXQUESTIONSPERGAME As Integer = 15
    Const intTIMEONQUESTION As Integer = 60            'the number of seconds on each question can be changed if needed(no negatives allowed)
    Const intMAXMONEY As Integer = 15000000             'the maximum amount of money the player is allowed to have
    Dim intMoneyWon As Integer = 0                      'The Amount of money the player has won
    Dim stcQuestionsDataBase(47) As QuestionsDataBase   'array for storing question not currently in use
    Dim stcQuestionsInUse(14) As QuestionsInUse         'array for storing questions in use
    'Dim strPlayerName As String
    Dim intQuestionPlayerIsOn As Integer                'The index value of the current question player is on in stcQuestionsInUse
    Dim blnLifeLineUsed As Boolean                      'Stores a boolean if the life line has been used or not
    Dim intTimeLeft As Integer = intTIMEONQUESTION      'the amount of time left in each question stored as a global static
    'Dim lblAllLabelsOnForm() As Label = {Me.lblOptionA, Me.lblOptionB, Me.lblOptionC, Me.lblOptionD, Me.lblTimer, Me.lblQuestion}
    'Dim btnAllButtonsOnForm() As Button = {Me.btnAnswerChoiceA, Me.btnAnswerChoiceB, Me.btnAnswerChoiceC, Me.btnAnswerChoiceD, Me.btnLifeLine, Me.btnWalkAway}

    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    'When the user clicks the start new game button or new game button in the program menu this procedure calls the required
    'procedures and sets the default settings
    '
    'post: BlnLifeLineUsed is set to false, intQuestionPlayerIsOn is set to 0, LoadQuestions is called to load the question 
    'data base into the array, RandomizeQuestions is called, LoadNextQuestion is called to load the next question in array.
    'SelectItem is called to select the correct index on the list box corresponding to the current question,
    'All labels and buttons exept btnStartNewGame is set visable to true including mnuPlayGame mnuWalkAway and mnuLifeLine
    'EndGameCelebration timer is enabled set to false, prgtimeLeft and prgTimeRight is visable is set to true, 
    'btnStartNewGame is visable is set to false
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Private Sub NewGame(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuNewGame.Click, btnStartNewGame.Click

        'Dim btnOptions() As Button = {Me.btnAnswerChoiceA, Me.btnAnswerChoiceB, Me.btnAnswerChoiceC, Me.btnAnswerChoiceD}
        Dim blnLifeLineUsed As Boolean
        Dim btnOptionsOne() As Button = {Me.btnAnswerChoiceA, Me.btnAnswerChoiceB, Me.btnAnswerChoiceC, Me.btnAnswerChoiceD}

        blnLifeLineUsed = False
        intQuestionPlayerIsOn = 0

        Call LoadQuestions(stcQuestionsDataBase)
        Call RandomizeQuestions(stcQuestionsInUse, stcQuestionsDataBase)
        Call LoadNextQuestion(stcQuestionsInUse, intQuestionPlayerIsOn, btnOptionsOne, Me.lblQuestion, _
                              intMoneyWon, Me.lblTimer, Me.tmrQuestionTimer, Me.picHintPicture)
        Call SelectItem(Me.lstMoneyWon, intQuestionPlayerIsOn)

        Me.lblOptionA.Visible = True
        Me.lblOptionB.Visible = True
        Me.lblOptionC.Visible = True
        Me.lblOptionD.Visible = True
        Me.btnAnswerChoiceA.Visible = True
        Me.btnAnswerChoiceB.Visible = True
        Me.btnAnswerChoiceC.Visible = True
        Me.btnAnswerChoiceD.Visible = True
        Me.mnuLifeLine.Visible = True
        Me.btnLifeLine.Enabled = True
        Me.mnuWalkAway.Visible = True
        Me.btnWalkAway.Visible = True
        Me.btnLifeLine.Visible = True
        Me.picHintPicture.Visible = True
        Me.lstMoneyWon.Visible = True
        Me.lblQuestion.Visible = True
        Me.lblTimer.Visible = True
        'Me.lstMoneyWon.Items.Add("Question #")
        ' Me.picHintPicture.Image = Image.FromFile("WhoWantsToBeAMillionare.png")
        Me.btnStartNewGame.Visible = False
        Me.prgTimeLeft.Visible = True
        Me.prgTimeRight.Visible = True
        Me.tmrEndGameCelebration.Enabled = False

    End Sub

    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    'When the user Clicks a btnAnswerChoice button depending on the tag the CheckAnswer procedure is called to
    'check if the answer is correct.
    '
    'post: when user clicks either btnAnswerChoiceA, btnAnswerChoiceB, btnAnswerChoiceC, btnAnswerChoiceD the 
    'tag is stored in chrPlayerAnswerChosen as a char (character) then it is passed to check answer to compare
    ' the text on the button to the correct answer, strAnswerClicked is the text on the button that the user 
    'clicked and that is compared to the right answer in check answer.
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Private Sub btnOption_Click(ByVal sender As Object, ByVal e As System.EventArgs) _
        Handles btnAnswerChoiceA.Click, btnAnswerChoiceB.Click, btnAnswerChoiceC.Click, btnAnswerChoiceD.Click

        Dim btnButtonClicked As Button = sender
        Dim chrPlayerAnswerChosen As Char
        Dim blnWrongAnswer As Boolean
        Dim strAnswerClicked As String
        Dim btnOptionsTwo() As Button = {Me.btnAnswerChoiceA, Me.btnAnswerChoiceB, Me.btnAnswerChoiceC, Me.btnAnswerChoiceD}
        Dim lblAllLabelsOnForm() As Label = {Me.lblOptionA, Me.lblOptionB, Me.lblOptionC, Me.lblOptionD, Me.lblTimer, Me.lblQuestion}
        Dim btnAllButtonsOnForm() As Button = {Me.btnAnswerChoiceA, Me.btnAnswerChoiceB, Me.btnAnswerChoiceC, Me.btnAnswerChoiceD, _
                                               Me.btnLifeLine, Me.btnWalkAway, Me.btnStartNewGame}

        chrPlayerAnswerChosen = btnButtonClicked.Tag

        Select Case chrPlayerAnswerChosen
            Case "a"
                strAnswerClicked = Me.btnAnswerChoiceA.Text
            Case "b"
                strAnswerClicked = Me.btnAnswerChoiceB.Text
            Case "c"
                strAnswerClicked = Me.btnAnswerChoiceC.Text
            Case "d"
                strAnswerClicked = Me.btnAnswerChoiceD.Text
        End Select

        Me.tmrQuestionTimer.Enabled = False

        Call CheckAnswer(strAnswerClicked, blnWrongAnswer, stcQuestionsInUse, btnOptionsTwo, intQuestionPlayerIsOn, _
                         chrPlayerAnswerChosen, Me.lblQuestion, Me.lstMoneyWon, Me.lblTimer, Me.tmrQuestionTimer, lblAllLabelsOnForm, _
                         btnAllButtonsOnForm, Me.mnuWalkAway, Me.mnuLifeLine, Me.picHintPicture, Me.prgTimeLeft, Me.prgTimeRight, Me.tmrEndGameCelebration)

    End Sub

    Private Sub mnuExit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuExit.Click

        End

    End Sub

    Private Sub mnuHelpAndInstructions_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuHelpAndInstructions.Click

        MessageBox.Show("Who wants to be a millionare was made by me (Abinash Singh), this game is based on the tv show. To play click new game" _
        & " in the program menu when your ready, choose the correct answer shown in the button, if the answer you chose is correct" _
        & " the button will turn green, else if it is wrong the button will turn red. If you wish to use the 50/50 life line, click the life line " _
        & "Button it will remove 2 of the wrong answers, by making them turn red. If you wish to walk away with the money and you have won more than" _
        & " $20,000 you can walk away and win, if you get to the top amount of $15000000 you win, but thats only if you get through all of the hardest" _
        & " questions ever asked by this.")

    End Sub

    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    'Loads all the listed questions answers and wrong answers into the array 
    '
    'post: the questions are loaded into the stcQuestions in use array
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Sub LoadQuestions(ByRef stcQuestionDataBase() As QuestionsDataBase)

        'stcQuestionDataBase(0).strQuestion = ""
        'stcQuestionDataBase(0).strCorrectAnswer = ""       The Format of listing the questions should be exactly like this example
        'stcBadAnswers(0).strWrongAnswerOne = ""
        'stcBadAnswers(0).strWrongAnswerTwo = ""
        'stcBadAnswers(0).strWrongAnswerThree = ""
        'stcQuestionDataBase(0).strHintPicture = ""         Stores the name of the picture file

        stcQuestionDataBase(0).strQuestion = "How Long Did The Rawandan Genocide Last?"
        stcQuestionDataBase(0).strCorrectAnswer = "100 Days"
        stcQuestionDataBase(0).strWrongAnswerOne = "300 Days"
        stcQuestionDataBase(0).strWrongAnswerTwo = "2 Years"
        stcQuestionDataBase(0).strWrongAnswerThree = "3 Weeks"
        stcQuestionDataBase(0).strHintPicture = ""

        stcQuestionDataBase(1).strQuestion = "How Many Atoms Are Required To Form A Covalent Bond?"
        stcQuestionDataBase(1).strCorrectAnswer = "2"
        stcQuestionDataBase(1).strWrongAnswerOne = "6"
        stcQuestionDataBase(1).strWrongAnswerTwo = "18"
        stcQuestionDataBase(1).strWrongAnswerThree = "3"
        stcQuestionDataBase(1).strHintPicture = ""

        stcQuestionDataBase(2).strQuestion = "Who Was The President During The Great Depression?"
        stcQuestionDataBase(2).strCorrectAnswer = "Herbert Clark Hoover"
        stcQuestionDataBase(2).strWrongAnswerOne = "Barack Obama"
        stcQuestionDataBase(2).strWrongAnswerTwo = "George Walker Bush"
        stcQuestionDataBase(2).strWrongAnswerThree = "Donald John Trump"
        stcQuestionDataBase(2).strHintPicture = ""

        stcQuestionDataBase(3).strQuestion = "Vehicles from which country use the international registration letters WG?"
        stcQuestionDataBase(3).strCorrectAnswer = "Grenada"
        stcQuestionDataBase(3).strWrongAnswerOne = "Canada"
        stcQuestionDataBase(3).strWrongAnswerTwo = "Germany"
        stcQuestionDataBase(3).strWrongAnswerThree = "United States of America"
        stcQuestionDataBase(3).strHintPicture = ""

        stcQuestionDataBase(4).strQuestion = "What was Elton John’s first US No 1 hit?"
        stcQuestionDataBase(4).strCorrectAnswer = "Crocodile Rock"
        stcQuestionDataBase(4).strWrongAnswerOne = "Pooch Jam"
        stcQuestionDataBase(4).strWrongAnswerTwo = "Beetle Rock"
        stcQuestionDataBase(4).strWrongAnswerThree = "Bodak Yellow"
        stcQuestionDataBase(4).strHintPicture = ""

        stcQuestionDataBase(5).strQuestion = "In which decade was the Oral Roberts University founded at Tulsa?"
        stcQuestionDataBase(5).strCorrectAnswer = "1960s"
        stcQuestionDataBase(5).strWrongAnswerOne = "1980s"
        stcQuestionDataBase(5).strWrongAnswerTwo = "1990s"
        stcQuestionDataBase(5).strWrongAnswerThree = "1960s"
        stcQuestionDataBase(5).strHintPicture = ""

        stcQuestionDataBase(6).strQuestion = "What was the profession of Louis Henry Sullivan?"
        stcQuestionDataBase(6).strCorrectAnswer = "Architect"
        stcQuestionDataBase(6).strWrongAnswerOne = "Artist"
        stcQuestionDataBase(6).strWrongAnswerTwo = "Engineer"
        stcQuestionDataBase(6).strWrongAnswerThree = "Clerk"
        stcQuestionDataBase(6).strHintPicture = ""

        stcQuestionDataBase(7).strQuestion = "Which space craft set off for Jupiter in 1972?"
        stcQuestionDataBase(7).strCorrectAnswer = "Pioneer 10"
        stcQuestionDataBase(7).strWrongAnswerOne = "Voyager 1"
        stcQuestionDataBase(7).strWrongAnswerTwo = "Soyuz"
        stcQuestionDataBase(7).strWrongAnswerThree = "Sputnik 1"
        stcQuestionDataBase(7).strHintPicture = ""

        stcQuestionDataBase(8).strQuestion = "What was the name of the Italian cruise ship hijacked by Palestinian terrorists in October 1985?"
        stcQuestionDataBase(8).strCorrectAnswer = "Achille Lauro"
        stcQuestionDataBase(8).strWrongAnswerOne = "Admiral, SS"
        stcQuestionDataBase(8).strWrongAnswerTwo = "Adriyatik, MS UND"
        stcQuestionDataBase(8).strWrongAnswerThree = "Aeolus"
        stcQuestionDataBase(8).strHintPicture = ""

        stcQuestionDataBase(9).strQuestion = "Bradley international airport is in which US state?"
        stcQuestionDataBase(9).strCorrectAnswer = "Connecticut"
        stcQuestionDataBase(9).strWrongAnswerOne = "Kentucky"
        stcQuestionDataBase(9).strWrongAnswerTwo = "Colarado"
        stcQuestionDataBase(9).strWrongAnswerThree = "New York"
        stcQuestionDataBase(9).strHintPicture = ""

        stcQuestionDataBase(10).strQuestion = "In which US state is John F Kennedy buried?"
        stcQuestionDataBase(10).strCorrectAnswer = "Virginia"
        stcQuestionDataBase(10).strWrongAnswerOne = "Texas"
        stcQuestionDataBase(10).strWrongAnswerTwo = "Kansas"
        stcQuestionDataBase(10).strWrongAnswerThree = "Idaho"
        stcQuestionDataBase(10).strHintPicture = ""

        stcQuestionDataBase(11).strQuestion = "What were UK’s Queen Elizabeth and Prince Philip given as a present for baby Prince Andrew while on a visit to the Gambia?"
        stcQuestionDataBase(11).strCorrectAnswer = "Baby Crocodile"
        stcQuestionDataBase(11).strWrongAnswerOne = "Panda"
        stcQuestionDataBase(11).strWrongAnswerTwo = "Bear"
        stcQuestionDataBase(11).strWrongAnswerThree = "Baby Turtle"
        stcQuestionDataBase(11).strHintPicture = ""

        stcQuestionDataBase(12).strQuestion = "Who was the first artist to enter the US album chart at No 1?"
        stcQuestionDataBase(12).strCorrectAnswer = "Elton John"
        stcQuestionDataBase(12).strWrongAnswerOne = "Bob Ross"
        stcQuestionDataBase(12).strWrongAnswerTwo = "Vincent van Gogh"
        stcQuestionDataBase(12).strWrongAnswerThree = "Leonardo da Vinci"
        stcQuestionDataBase(12).strHintPicture = ""

        stcQuestionDataBase(13).strQuestion = "What are the international registration letters of a vehicle from Guyana?"
        stcQuestionDataBase(13).strCorrectAnswer = "GUY"
        stcQuestionDataBase(13).strWrongAnswerOne = "AFG"
        stcQuestionDataBase(13).strWrongAnswerTwo = "AND"
        stcQuestionDataBase(13).strWrongAnswerThree = "BUR"
        stcQuestionDataBase(13).strHintPicture = ""

        stcQuestionDataBase(14).strQuestion = "Which US No 1 single came from Diana Ross’s platinum album Diana?"
        stcQuestionDataBase(14).strCorrectAnswer = "Upside Down"
        stcQuestionDataBase(14).strWrongAnswerOne = "Down Side Up"
        stcQuestionDataBase(14).strWrongAnswerTwo = "Right Side Up"
        stcQuestionDataBase(14).strWrongAnswerThree = "Im Comming Out"
        stcQuestionDataBase(14).strHintPicture = ""

        stcQuestionDataBase(15).strQuestion = "What was the title of Kitty Kelley’s book about Elizabeth Taylor?"
        stcQuestionDataBase(15).strCorrectAnswer = "Elizabeth Taylor: The Last Star"
        stcQuestionDataBase(15).strWrongAnswerOne = "Elizabeth Taylor: The Lost Star"
        stcQuestionDataBase(15).strWrongAnswerTwo = "Elizabeth Taylor: The Last Moon"
        stcQuestionDataBase(15).strWrongAnswerThree = "Elizabeth Taylor: The First"
        stcQuestionDataBase(15).strHintPicture = ""

        stcQuestionDataBase(16).strQuestion = "What was Eddie Fisher’s last top ten hit?"
        stcQuestionDataBase(16).strCorrectAnswer = "Cindy Oh Cindy"
        stcQuestionDataBase(16).strWrongAnswerOne = "Cindy On Cindy"
        stcQuestionDataBase(16).strWrongAnswerTwo = " Cindy Oh Cindley"
        stcQuestionDataBase(16).strWrongAnswerThree = " Cindy Oh Oh Cindy"
        stcQuestionDataBase(16).strHintPicture = ""

        stcQuestionDataBase(17).strQuestion = "Which European newspaper published detailed photographs of the car crash involving Diana, Princess of Wales?"
        stcQuestionDataBase(17).strCorrectAnswer = "Bild"
        stcQuestionDataBase(17).strWrongAnswerOne = "New York Times"
        stcQuestionDataBase(17).strWrongAnswerTwo = "The Wall Street Journal"
        stcQuestionDataBase(17).strWrongAnswerThree = "Los Angeles Times"
        stcQuestionDataBase(17).strHintPicture = ""

        stcQuestionDataBase(18).strQuestion = "Which country does the airline Sansa come from?"
        stcQuestionDataBase(18).strCorrectAnswer = "Costa Rica"
        stcQuestionDataBase(18).strWrongAnswerOne = "Columbia"
        stcQuestionDataBase(18).strWrongAnswerTwo = "Brazil"
        stcQuestionDataBase(18).strWrongAnswerThree = "Cuba"
        stcQuestionDataBase(18).strHintPicture = ""

        stcQuestionDataBase(19).strQuestion = "In which English city was Cary Grant born?"
        stcQuestionDataBase(19).strCorrectAnswer = "Bristol"
        stcQuestionDataBase(19).strWrongAnswerOne = "Montreal"
        stcQuestionDataBase(19).strWrongAnswerTwo = "Weston-super-Mare"
        stcQuestionDataBase(19).strWrongAnswerThree = "Wells"
        stcQuestionDataBase(19).strHintPicture = ""

        stcQuestionDataBase(20).strQuestion = "Which actress was born on exactly the same day as Al Gore?"
        stcQuestionDataBase(20).strCorrectAnswer = "Rhea Pearlman"
        stcQuestionDataBase(20).strWrongAnswerOne = "Brittney Spears"
        stcQuestionDataBase(20).strWrongAnswerTwo = "Marty Mcfly"
        stcQuestionDataBase(20).strWrongAnswerThree = "Cardi B"
        stcQuestionDataBase(20).strHintPicture = ""

        stcQuestionDataBase(20).strQuestion = "What is Mel Gibson’s middle name?"
        stcQuestionDataBase(20).strCorrectAnswer = "Columcille"
        stcQuestionDataBase(20).strWrongAnswerOne = "Cucumber"
        stcQuestionDataBase(20).strWrongAnswerTwo = "Westing"
        stcQuestionDataBase(20).strWrongAnswerThree = "Wells"
        stcQuestionDataBase(20).strHintPicture = ""

        stcQuestionDataBase(21).strQuestion = "In which country was Julie Christie born?"
        stcQuestionDataBase(21).strCorrectAnswer = "India"
        stcQuestionDataBase(21).strWrongAnswerOne = "Canada"
        stcQuestionDataBase(21).strWrongAnswerTwo = "Britain"
        stcQuestionDataBase(21).strWrongAnswerThree = "Afghanistan"
        stcQuestionDataBase(21).strHintPicture = ""

        stcQuestionDataBase(22).strQuestion = "What type of material did Cupples and Leon publish?"
        stcQuestionDataBase(22).strCorrectAnswer = "Comic books"
        stcQuestionDataBase(22).strWrongAnswerOne = "Graphic Novels"
        stcQuestionDataBase(22).strWrongAnswerTwo = "Magazines"
        stcQuestionDataBase(22).strWrongAnswerThree = "Newspapers"
        stcQuestionDataBase(22).strHintPicture = ""

        stcQuestionDataBase(23).strQuestion = "Freeport international airport is in which country?"
        stcQuestionDataBase(23).strCorrectAnswer = "The Bahamas"
        stcQuestionDataBase(23).strWrongAnswerOne = "Cuba"
        stcQuestionDataBase(23).strWrongAnswerTwo = "Caribbean"
        stcQuestionDataBase(23).strWrongAnswerThree = "Saudi Arabia"
        stcQuestionDataBase(23).strHintPicture = ""

        stcQuestionDataBase(24).strQuestion = "What was the profession of Dorothea Lange?"
        stcQuestionDataBase(24).strCorrectAnswer = "Photographer"
        stcQuestionDataBase(24).strWrongAnswerOne = "Doctor"
        stcQuestionDataBase(24).strWrongAnswerTwo = "Veterinarian"
        stcQuestionDataBase(24).strWrongAnswerThree = "Programmer"
        stcQuestionDataBase(24).strHintPicture = ""

        stcQuestionDataBase(25).strQuestion = "Sculptor Drederic Auguste Bartholdi based the face of the Statue of Liberty on whom?"
        stcQuestionDataBase(25).strCorrectAnswer = "His Mother"
        stcQuestionDataBase(25).strWrongAnswerOne = "His Father"
        stcQuestionDataBase(25).strWrongAnswerTwo = "His Brother"
        stcQuestionDataBase(25).strWrongAnswerThree = "His Sister"
        stcQuestionDataBase(25).strHintPicture = ""

        stcQuestionDataBase(26).strQuestion = "Which group made the albums Bare Trees and Penguin? "
        stcQuestionDataBase(26).strCorrectAnswer = "Fleetwood Mac"
        stcQuestionDataBase(26).strWrongAnswerOne = "The Bettles"
        stcQuestionDataBase(26).strWrongAnswerTwo = "Eagles"
        stcQuestionDataBase(26).strWrongAnswerThree = "Led Zeppelin"
        stcQuestionDataBase(26).strHintPicture = ""

        stcQuestionDataBase(27).strQuestion = "What color were the covers of the crime novels published in  the 1930s by Victor Gollancz?"
        stcQuestionDataBase(27).strCorrectAnswer = "Yellow"
        stcQuestionDataBase(27).strWrongAnswerOne = "Red with Blood"
        stcQuestionDataBase(27).strWrongAnswerTwo = "Green"
        stcQuestionDataBase(27).strWrongAnswerThree = "Orange"
        stcQuestionDataBase(27).strHintPicture = ""

        stcQuestionDataBase(28).strQuestion = "Where in the UK was Robert Carlyle born?"
        stcQuestionDataBase(28).strCorrectAnswer = "Glasgow"
        stcQuestionDataBase(28).strWrongAnswerOne = "Manchester"
        stcQuestionDataBase(28).strWrongAnswerTwo = "London"
        stcQuestionDataBase(28).strWrongAnswerThree = "Leeds"
        stcQuestionDataBase(28).strHintPicture = "Clydeside1.jpg" 'https://www.google.ca/imgres?imgurl=https://peoplemakeglasgow.com/images/Clydeside1.jpg&imgrefurl=https://peoplemakeglasgow.com/&h=500&w=995&tbnid=2NPL0J4WAD-MKM:&tbnh=159&tbnw=317&usg=__JHRQf_at9ijY5J0floSutxIVxyM%3D&vet=10ahUKEwiZzo2bpMvYAhVk34MKHYeXC6cQ_B0InAEwEQ..i&docid=Drw2EJnOGsS6pM&itg=1&sa=X&ved=0ahUKEwiZzo2bpMvYAhVk34MKHYeXC6cQ_B0InAEwEQ

        stcQuestionDataBase(29).strQuestion = "Which City Is This?"
        stcQuestionDataBase(29).strCorrectAnswer = "Sydney"
        stcQuestionDataBase(29).strWrongAnswerOne = "New York"
        stcQuestionDataBase(29).strWrongAnswerTwo = "Naples"
        stcQuestionDataBase(29).strWrongAnswerThree = "Buenos Aires"
        stcQuestionDataBase(29).strHintPicture = "sydney.jpg" 'https://lh4.googleusercontent.com/proxy/LooC8mhKLCm0UobLNa9yijxOna_Nc6En1GJmku9MMW9UJls8lJPJo0LU20OddjD0PnL_p_6BVAeT3hAYT-nXIV0WvEU06PMKfgTn8qq4_4igUMVv0oD71LpQze5qjn3to36s8I8ph47AcHoE7UeBVCW_j7Ao8Q=w100-h134-n-k-no

        stcQuestionDataBase(30).strQuestion = "Which brothers were Warner Bros’ first major record success?"
        stcQuestionDataBase(30).strCorrectAnswer = "Everly Brothers"
        stcQuestionDataBase(30).strWrongAnswerOne = "Maggie and Jake Gyllenhaal"
        stcQuestionDataBase(30).strWrongAnswerTwo = "James and Dave Franco"
        stcQuestionDataBase(30).strWrongAnswerThree = "James and Solange Knowles"
        stcQuestionDataBase(30).strHintPicture = "the-everly-brothers.jpg" 'https://www.google.ca/imgres?imgurl=https://www.biography.com/.image/t_share/MTE4MDAzNDEwMTUwMTMxMjE0/the-everly-brothers.jpg&imgrefurl=https://www.biography.com/people/groups/the-everly-brothers&h=935&w=1200&tbnid=z4BYF2-dECjpIM:&tbnh=156&tbnw=199&usg=__3buo-Co5DBqqhP4M5AyUqlUaxVA%3D&vet=10ahUKEwjNzIy-qMvYAhUow4MKHYQQAkQQ_B0IoAEwEw..i&docid=kO-DxC1J7rrR2M&itg=1&sa=X&ved=0ahUKEwjNzIy-qMvYAhUow4MKHYQQAkQQ_B0IoAEwEw

        stcQuestionDataBase(31).strQuestion = "Where was Che Guevara killed?"
        stcQuestionDataBase(31).strCorrectAnswer = "Bolivia"
        stcQuestionDataBase(31).strWrongAnswerOne = "Cuba"
        stcQuestionDataBase(31).strWrongAnswerTwo = "Hawaii"
        stcQuestionDataBase(31).strWrongAnswerThree = "Mexico"
        stcQuestionDataBase(31).strHintPicture = "che-prisoner.jpg" 'https://media.newyorker.com/photos/590959c02179605b11ad4856/master/w_727,c_limit/che-prisoner.jpg

        stcQuestionDataBase(32).strQuestion = "What is Whoopi Goldberg’s real name?"
        stcQuestionDataBase(32).strCorrectAnswer = "Caryn Johnson"
        stcQuestionDataBase(32).strWrongAnswerOne = "Harvey Johnson"
        stcQuestionDataBase(32).strWrongAnswerTwo = "Dave Johnson"
        stcQuestionDataBase(32).strWrongAnswerThree = "Kim Johnson"
        stcQuestionDataBase(32).strHintPicture = "1200px-Whoopi_Goldberg_at_a_NYC_No_on_Proposition_8_Rally.jpg" 'https://www.google.ca/imgres?imgurl=https://upload.wikimedia.org/wikipedia/commons/thumb/4/4b/Whoopi_Goldberg_at_a_NYC_No_on_Proposition_8_Rally.jpg/1200px-Whoopi_Goldberg_at_a_NYC_No_on_Proposition_8_Rally.jpg&imgrefurl=https://sw.wikipedia.org/wiki/Whoopi_Goldberg&h=1555&w=1200&tbnid=KPdHBG-wUIpRyM:&tbnh=186&tbnw=143&usg=__1v6d0B7LiNJ1yPrW4yJkE6UFKZc%3D&vet=10ahUKEwiWwtj8qcvYAhUN8YMKHeUrAKwQ_B0IngEwEw..i&docid=WMxSkrISYK3Z4M&itg=1&sa=X&ved=0ahUKEwiWwtj8qcvYAhUN8YMKHeUrAKwQ_B0IngEwEw

        stcQuestionDataBase(33).strQuestion = "Ralph Craig ran the 100m for the US in 1912; when did he next compete in the Olympics?"
        stcQuestionDataBase(33).strCorrectAnswer = "1948"
        stcQuestionDataBase(33).strWrongAnswerOne = "1914"
        stcQuestionDataBase(33).strWrongAnswerTwo = "1921"
        stcQuestionDataBase(33).strWrongAnswerThree = "2016"
        stcQuestionDataBase(33).strHintPicture = "21-iunie-1889-s-a-nascut-ralph-craig-campion-olimpic-la-100-si-200-de-metri-in-1912-39547.jpg" 'https://www.google.ca/search?q=Ralph+Craig&safe=strict&rlz=1C1GCEA_enCA762CA762&source=lnms&tbm=isch&sa=X&ved=0ahUKEwjX1Nqzq8vYAhUD34MKHel3B0QQ_AUICigB&biw=958&bih=890#imgrc=0xVJ221bSMC-VM:

        stcQuestionDataBase(34).strQuestion = "In which year was the Tampa-based University of South Florida founded?"
        stcQuestionDataBase(34).strCorrectAnswer = "1956"
        stcQuestionDataBase(34).strWrongAnswerOne = "1996"
        stcQuestionDataBase(34).strWrongAnswerTwo = "2001"
        stcQuestionDataBase(34).strWrongAnswerThree = "1808"
        stcQuestionDataBase(34).strHintPicture = "USFSP-FB.jpg" 'https://www.google.ca/search?q=University+of+South+Florida&safe=strict&rlz=1C1GCEA_enCA762CA762&source=lnms&tbm=isch&sa=X&ved=0ahUKEwiykIGFrMvYAhUC2oMKHfABCqkQ_AUICygC&biw=958&bih=890#imgrc=finZgnnBkilNQM:

        stcQuestionDataBase(35).strQuestion = "Which writer said,'Where large sums of money are involved, it is advisable to trust nobody?'"
        stcQuestionDataBase(35).strCorrectAnswer = "Agatha Christie"
        stcQuestionDataBase(35).strWrongAnswerOne = "Ernest Hemingway"
        stcQuestionDataBase(35).strWrongAnswerTwo = "Mark Twain"
        stcQuestionDataBase(35).strWrongAnswerThree = "Stephen King"
        stcQuestionDataBase(35).strHintPicture = "agatha-christie-9247405-1-402.jpg" 'https://www.google.ca/imgres?imgurl=https://www.biography.com/.image/t_share/MTE5NDg0MDU0OTI0MDY4MzY3/agatha-christie-9247405-1-402.jpg&imgrefurl=https://www.biography.com/people/agatha-christie-9247405&h=1200&w=1200&tbnid=hhAc5yrR4mVTQM:&tbnh=186&tbnw=186&usg=__Xk1xiRv-LzUKoOn6Sfumhtnlvw4%3D&vet=10ahUKEwj-557PrMvYAhVK7IMKHTtzCqkQ_B0IswEwFw..i&docid=FkyeMMGtd4c_VM&itg=1&sa=X&ved=0ahUKEwj-557PrMvYAhVK7IMKHTtzCqkQ_B0IswEwFw

        stcQuestionDataBase(36).strQuestion = "What was Richard Gere’s first movie?"
        stcQuestionDataBase(36).strCorrectAnswer = "Report To The Commissioner"
        stcQuestionDataBase(36).strWrongAnswerOne = "Pretty Woman"
        stcQuestionDataBase(36).strWrongAnswerTwo = "An Officer an a Gentleman"
        stcQuestionDataBase(36).strWrongAnswerThree = "American Giagolo"
        stcQuestionDataBase(36).strHintPicture = "Richard gere.jpg" 'https://www.google.ca/search?q=Richard+Gere%E2%80%99s&safe=strict&rlz=1C1GCEA_enCA762CA762&tbm=isch&source=iu&ictx=1&fir=rjZk3hl1a_O5bM%253A%252CUB14XzAzhiSfVM%252C_&usg=__1FHsMD7iQvuArvp202lVlfXzLMw%3D&sa=X&ved=0ahUKEwjC7aKsrcvYAhUI9IMKHfJvCqYQ_h0IrgEwFw#imgrc=wwx2601YHNLlbM:

        stcQuestionDataBase(37).strQuestion = "In which year did Frank Zappa die?"
        stcQuestionDataBase(37).strCorrectAnswer = "1993"
        stcQuestionDataBase(37).strWrongAnswerOne = "1994"
        stcQuestionDataBase(37).strWrongAnswerTwo = "1808"
        stcQuestionDataBase(37).strWrongAnswerThree = "1666"
        stcQuestionDataBase(37).strHintPicture = "FRANK ZAPPA.jpg" 'https://www.google.ca/imgres?imgurl=https://cps-static.rovicorp.com/3/JPG_400/MI0003/389/MI0003389043.jpg?partner%3Dallrovi.com&imgrefurl=https://www.allmusic.com/artist/frank-zappa-mn0000138699&h=411&w=400&tbnid=SpBVqP3_C3LotM:&tbnh=186&tbnw=181&usg=__0FjlRAlZGKIDnexm14sE9a6KBgw%3D&vet=10ahUKEwjjjf6xrsvYAhWkzIMKHRTRA6oQ_B0IrAEwFQ..i&docid=S4z2H6CXhWqZbM&itg=1&sa=X&ved=0ahUKEwjjjf6xrsvYAhWkzIMKHRTRA6oQ_B0IrAEwFQ

        stcQuestionDataBase(38).strQuestion = "Vehicles from which county use the international registration letter LS?"
        stcQuestionDataBase(38).strCorrectAnswer = "Lesotho"
        stcQuestionDataBase(38).strWrongAnswerOne = "Lativa"
        stcQuestionDataBase(38).strWrongAnswerTwo = "Lithuania"
        stcQuestionDataBase(38).strWrongAnswerThree = "Liberia"
        stcQuestionDataBase(38).strHintPicture = "1200px-Flag_of_Lesotho.svg.png" 'https://www.google.ca/search?q=Lesotho&safe=strict&rlz=1C1GCEA_enCA762CA762&source=lnms&tbm=isch&sa=X&ved=0ahUKEwi6xIjersvYAhUr6IMKHdSCC6cQ_AUIDCgD&biw=958&bih=890#imgrc=tHxKoFDNbkzBXM:

        stcQuestionDataBase(39).strQuestion = "In which decade was the American Ballet set up in New York?"
        stcQuestionDataBase(39).strCorrectAnswer = "1933"
        stcQuestionDataBase(39).strWrongAnswerOne = "1944"
        stcQuestionDataBase(39).strWrongAnswerTwo = "1988"
        stcQuestionDataBase(39).strWrongAnswerThree = "1966"
        stcQuestionDataBase(39).strHintPicture = ""

        stcQuestionDataBase(40).strQuestion = "Who is this?"
        stcQuestionDataBase(40).strCorrectAnswer = "cris"
        stcQuestionDataBase(40).strWrongAnswerOne = "Cris"
        stcQuestionDataBase(40).strWrongAnswerTwo = "Kris"
        stcQuestionDataBase(40).strWrongAnswerThree = "Kriz"
        stcQuestionDataBase(40).strHintPicture = "cris.bmp"

        stcQuestionDataBase(41).strQuestion = "Douala international airport is in which country?"
        stcQuestionDataBase(41).strCorrectAnswer = "Cameroon"
        stcQuestionDataBase(41).strWrongAnswerOne = "Syria"
        stcQuestionDataBase(41).strWrongAnswerTwo = "Cambodia"
        stcQuestionDataBase(41).strWrongAnswerThree = "Uganda"
        stcQuestionDataBase(41).strHintPicture = "1200px-Flag_of_Cameroon.svg.png" 'https://www.google.ca/search?safe=strict&rlz=1C1GCEA_enCA762CA762&biw=958&bih=939&tbm=isch&sa=1&ei=hiZeWt6RCuzPjwS-8IiwDg&q=cameroon&oq=cameroon&gs_l=psy-ab.3..0l10.2184.96669.0.96809.17.12.3.1.1.0.120.976.10j1.11.0....0...1c.1.64.psy-ab..3.14.931...0i5i30k1j0i13i5i30k1j0i24k1j0i67k1.0.oDApXCj0qpo#imgrc=Q2b4kpvPw0zidM:

        stcQuestionDataBase(42).strQuestion = "Who won ESL One Cologne in 2017"
        stcQuestionDataBase(42).strCorrectAnswer = "SK Gaming"
        stcQuestionDataBase(42).strWrongAnswerOne = "Cloud9"
        stcQuestionDataBase(42).strWrongAnswerTwo = "Faze"
        stcQuestionDataBase(42).strWrongAnswerThree = "NAV"
        stcQuestionDataBase(42).strHintPicture = ""

        stcQuestionDataBase(43).strQuestion = "At the beginning of the 1990s which country had most camels?"
        stcQuestionDataBase(43).strCorrectAnswer = "Somalia"
        stcQuestionDataBase(43).strWrongAnswerOne = "Bolivia"
        stcQuestionDataBase(43).strWrongAnswerTwo = "UAE"
        stcQuestionDataBase(43).strWrongAnswerThree = "Afganistan"
        stcQuestionDataBase(43).strHintPicture = ""

        stcQuestionDataBase(44).strQuestion = "In which American state are the Merril Collection and the Burke Museum of Fine Arts?"
        stcQuestionDataBase(44).strCorrectAnswer = "Texas"
        stcQuestionDataBase(44).strWrongAnswerOne = "Kansas"
        stcQuestionDataBase(44).strWrongAnswerTwo = "New York"
        stcQuestionDataBase(44).strWrongAnswerThree = "Washington"
        stcQuestionDataBase(44).strHintPicture = ""

        stcQuestionDataBase(45).strQuestion = "Which actor paid $93,500 for the baseball which rolled between Bill Buckner’s legs in game six of the 1986 World Series?"
        stcQuestionDataBase(45).strCorrectAnswer = "Charlie Sheen"
        stcQuestionDataBase(45).strWrongAnswerOne = "Denise Richards"
        stcQuestionDataBase(45).strWrongAnswerTwo = "Martin Sheen"
        stcQuestionDataBase(45).strWrongAnswerThree = "Brooke Mueller"
        stcQuestionDataBase(45).strHintPicture = ""

        stcQuestionDataBase(46).strQuestion = "Who was Theodore Roosevelt’s Vice President between 1905 and 1909?"
        stcQuestionDataBase(46).strCorrectAnswer = "Charles W. Fairbanks"
        stcQuestionDataBase(46).strWrongAnswerOne = "James S. Sherman"
        stcQuestionDataBase(46).strWrongAnswerTwo = "Garret Hobart"
        stcQuestionDataBase(46).strWrongAnswerThree = "Nicholas M. Butler"
        stcQuestionDataBase(46).strHintPicture = ""

        stcQuestionDataBase(47).strQuestion = "Which nation was the first to ratify the United Nations charter in 1945?"
        stcQuestionDataBase(47).strCorrectAnswer = "Nicaragua"
        stcQuestionDataBase(47).strWrongAnswerOne = "Costa Rica"
        stcQuestionDataBase(47).strWrongAnswerTwo = "Honduras"
        stcQuestionDataBase(47).strWrongAnswerThree = "Guatemala"
        stcQuestionDataBase(47).strHintPicture = ""

    End Sub

    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    'Randomizes the order the questions, the answers and the hintPicture, are loaded into the array stcQuestionsInUse
    '
    'post: The do loop generates fifteen intergers completely different from each other depending on the length
    'of the stcQuestionsdataBase array then the index of the elements corresponding to the stcQuestionsDataBase are loaded
    'into the stcQuestionsInUse array
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Sub RandomizeQuestions(ByRef stcQuestionsInUse() As QuestionsInUse, ByRef stcQuestionDataBase() As QuestionsDataBase)

        Dim intRandomIntegerOne As Integer
        Dim intRandomIntegerTwo As Integer
        Dim intRandomIntegerThree As Integer
        Dim intRandomIntegerFour As Integer
        Dim intRandomIntegerFive As Integer
        Dim intRandomIntegerSix As Integer
        Dim intRandomIntegerSeven As Integer
        Dim intRandomIntegerEight As Integer
        Dim intRandomIntegerNine As Integer
        Dim intRandomIntegerTen As Integer
        Dim intRandomIntegerEleven As Integer
        Dim intRandomIntegerTwelve As Integer
        Dim intRandomIntegerThirteen As Integer
        Dim intRandomIntegerFourteen As Integer
        Dim intRandomIntegerFifteen As Integer

        Do
            intRandomIntegerOne = RndInt(stcQuestionDataBase.Length - 1, 0)
            intRandomIntegerTwo = RndInt(stcQuestionDataBase.Length - 1, 0)
            intRandomIntegerThree = RndInt(stcQuestionDataBase.Length - 1, 0)
            intRandomIntegerFour = RndInt(stcQuestionDataBase.Length - 1, 0)
            intRandomIntegerFive = RndInt(stcQuestionDataBase.Length - 1, 0)
            intRandomIntegerSix = RndInt(stcQuestionDataBase.Length - 1, 0)
            intRandomIntegerSeven = RndInt(stcQuestionDataBase.Length - 1, 0)
            intRandomIntegerEight = RndInt(stcQuestionDataBase.Length - 1, 0)
            intRandomIntegerNine = RndInt(stcQuestionDataBase.Length - 1, 0)
            intRandomIntegerTen = RndInt(stcQuestionDataBase.Length - 1, 0)
            intRandomIntegerEleven = RndInt(stcQuestionDataBase.Length - 1, 0)
            intRandomIntegerTwelve = RndInt(stcQuestionDataBase.Length - 1, 0)
            intRandomIntegerThirteen = RndInt(stcQuestionDataBase.Length - 1, 0)
            intRandomIntegerFourteen = RndInt(stcQuestionDataBase.Length - 1, 0)
            intRandomIntegerFifteen = RndInt(stcQuestionDataBase.Length - 1, 0)
        Loop Until intRandomIntegerOne <> intRandomIntegerTwo And intRandomIntegerOne <> intRandomIntegerThree _
And intRandomIntegerOne <> intRandomIntegerFour And intRandomIntegerOne <> intRandomIntegerFive _
And intRandomIntegerOne <> intRandomIntegerSix And intRandomIntegerOne <> intRandomIntegerSeven _
And intRandomIntegerOne <> intRandomIntegerEight _
And intRandomIntegerOne <> intRandomIntegerNine And intRandomIntegerOne <> intRandomIntegerTen _
And intRandomIntegerOne <> intRandomIntegerEleven And intRandomIntegerOne <> intRandomIntegerTwelve _
And intRandomIntegerOne <> intRandomIntegerThirteen And intRandomIntegerOne <> intRandomIntegerFourteen _
And intRandomIntegerOne <> intRandomIntegerFifteen _
And intRandomIntegerTwo <> intRandomIntegerOne And intRandomIntegerTwo <> intRandomIntegerThree _
And intRandomIntegerTwo <> intRandomIntegerFour And intRandomIntegerTwo <> intRandomIntegerFive _
And intRandomIntegerTwo <> intRandomIntegerSix And intRandomIntegerTwo <> intRandomIntegerSeven _
And intRandomIntegerTwo <> intRandomIntegerEight _
And intRandomIntegerTwo <> intRandomIntegerNine And intRandomIntegerTwo <> intRandomIntegerTen _
And intRandomIntegerTwo <> intRandomIntegerEleven And intRandomIntegerTwo <> intRandomIntegerTwelve _
And intRandomIntegerTwo <> intRandomIntegerThirteen And intRandomIntegerTwo <> intRandomIntegerFourteen _
And intRandomIntegerTwo <> intRandomIntegerFifteen _
And intRandomIntegerThree <> intRandomIntegerTwo _
And intRandomIntegerThree <> intRandomIntegerFour And intRandomIntegerThree <> intRandomIntegerFive _
And intRandomIntegerThree <> intRandomIntegerSix And intRandomIntegerThree <> intRandomIntegerSeven _
And intRandomIntegerThree <> intRandomIntegerEight _
And intRandomIntegerThree <> intRandomIntegerNine And intRandomIntegerThree <> intRandomIntegerTen _
And intRandomIntegerThree <> intRandomIntegerEleven And intRandomIntegerThree <> intRandomIntegerTwelve _
And intRandomIntegerThree <> intRandomIntegerThirteen And intRandomIntegerThree <> intRandomIntegerFourteen _
And intRandomIntegerThree <> intRandomIntegerFifteen _
And intRandomIntegerFour <> intRandomIntegerOne And intRandomIntegerFour <> intRandomIntegerTwo _
And intRandomIntegerFour <> intRandomIntegerThree _
And intRandomIntegerFour <> intRandomIntegerFive _
And intRandomIntegerFour <> intRandomIntegerSix And intRandomIntegerFour <> intRandomIntegerSeven _
And intRandomIntegerFour <> intRandomIntegerEight _
And intRandomIntegerFour <> intRandomIntegerNine And intRandomIntegerFour <> intRandomIntegerTen _
And intRandomIntegerFour <> intRandomIntegerEleven And intRandomIntegerFour <> intRandomIntegerTwelve _
And intRandomIntegerFour <> intRandomIntegerThirteen And intRandomIntegerFour <> intRandomIntegerFourteen _
And intRandomIntegerFour <> intRandomIntegerFifteen _
And intRandomIntegerFive <> intRandomIntegerOne And intRandomIntegerFive <> intRandomIntegerTwo _
And intRandomIntegerFive <> intRandomIntegerThree _
And intRandomIntegerFive <> intRandomIntegerFour _
And intRandomIntegerFive <> intRandomIntegerSix And intRandomIntegerFive <> intRandomIntegerSeven _
And intRandomIntegerFive <> intRandomIntegerEight _
And intRandomIntegerFive <> intRandomIntegerNine And intRandomIntegerFive <> intRandomIntegerTen _
And intRandomIntegerFive <> intRandomIntegerEleven And intRandomIntegerFive <> intRandomIntegerTwelve _
And intRandomIntegerFive <> intRandomIntegerThirteen And intRandomIntegerFive <> intRandomIntegerFourteen _
And intRandomIntegerFive <> intRandomIntegerFifteen _
And intRandomIntegerSix <> intRandomIntegerOne And intRandomIntegerSix <> intRandomIntegerTwo _
And intRandomIntegerSix <> intRandomIntegerThree _
And intRandomIntegerSix <> intRandomIntegerFour And intRandomIntegerSix <> intRandomIntegerFive _
And intRandomIntegerSix <> intRandomIntegerSeven _
And intRandomIntegerSix <> intRandomIntegerEight _
And intRandomIntegerSix <> intRandomIntegerNine And intRandomIntegerSix <> intRandomIntegerTen _
And intRandomIntegerSix <> intRandomIntegerEleven And intRandomIntegerSix <> intRandomIntegerTwelve _
And intRandomIntegerSix <> intRandomIntegerThirteen And intRandomIntegerSix <> intRandomIntegerFourteen _
And intRandomIntegerSix <> intRandomIntegerFifteen _
And intRandomIntegerSeven <> intRandomIntegerOne And intRandomIntegerSeven <> intRandomIntegerTwo _
And intRandomIntegerSeven <> intRandomIntegerThree _
And intRandomIntegerSeven <> intRandomIntegerFour And intRandomIntegerSeven <> intRandomIntegerFive _
And intRandomIntegerSeven <> intRandomIntegerSix _
And intRandomIntegerSeven <> intRandomIntegerEight _
And intRandomIntegerSeven <> intRandomIntegerNine And intRandomIntegerSeven <> intRandomIntegerTen _
And intRandomIntegerSeven <> intRandomIntegerEleven And intRandomIntegerSeven <> intRandomIntegerTwelve _
And intRandomIntegerSeven <> intRandomIntegerThirteen And intRandomIntegerSeven <> intRandomIntegerFourteen _
And intRandomIntegerSeven <> intRandomIntegerFifteen _
And intRandomIntegerEight <> intRandomIntegerOne And intRandomIntegerEight <> intRandomIntegerTwo _
And intRandomIntegerEight <> intRandomIntegerThree _
And intRandomIntegerEight <> intRandomIntegerFour And intRandomIntegerEight <> intRandomIntegerFive _
And intRandomIntegerEight <> intRandomIntegerSix And intRandomIntegerEight <> intRandomIntegerSeven _
And intRandomIntegerEight <> intRandomIntegerNine And intRandomIntegerEight <> intRandomIntegerTen _
And intRandomIntegerEight <> intRandomIntegerEleven And intRandomIntegerEight <> intRandomIntegerTwelve _
And intRandomIntegerEight <> intRandomIntegerThirteen And intRandomIntegerEight <> intRandomIntegerFourteen _
And intRandomIntegerEight <> intRandomIntegerFifteen _
And intRandomIntegerNine <> intRandomIntegerOne And intRandomIntegerNine <> intRandomIntegerTwo _
And intRandomIntegerNine <> intRandomIntegerThree _
And intRandomIntegerNine <> intRandomIntegerFour And intRandomIntegerNine <> intRandomIntegerFive _
And intRandomIntegerNine <> intRandomIntegerSix And intRandomIntegerNine <> intRandomIntegerSeven _
And intRandomIntegerNine <> intRandomIntegerEight _
And intRandomIntegerNine <> intRandomIntegerTen _
And intRandomIntegerNine <> intRandomIntegerEleven And intRandomIntegerNine <> intRandomIntegerTwelve _
And intRandomIntegerNine <> intRandomIntegerThirteen And intRandomIntegerNine <> intRandomIntegerFourteen _
And intRandomIntegerNine <> intRandomIntegerFifteen _
And intRandomIntegerTen <> intRandomIntegerOne And intRandomIntegerTen <> intRandomIntegerTwo _
And intRandomIntegerTen <> intRandomIntegerThree _
And intRandomIntegerTen <> intRandomIntegerFour And intRandomIntegerTen <> intRandomIntegerFive _
And intRandomIntegerTen <> intRandomIntegerSix And intRandomIntegerTen <> intRandomIntegerSeven _
And intRandomIntegerTen <> intRandomIntegerEight _
And intRandomIntegerTen <> intRandomIntegerNine _
And intRandomIntegerTen <> intRandomIntegerEleven And intRandomIntegerTen <> intRandomIntegerTwelve _
And intRandomIntegerTen <> intRandomIntegerThirteen And intRandomIntegerTen <> intRandomIntegerFourteen _
And intRandomIntegerTen <> intRandomIntegerFifteen _
And intRandomIntegerEleven <> intRandomIntegerOne And intRandomIntegerEleven <> intRandomIntegerTwo _
And intRandomIntegerEleven <> intRandomIntegerThree _
And intRandomIntegerEleven <> intRandomIntegerFour And intRandomIntegerEleven <> intRandomIntegerFive _
And intRandomIntegerEleven <> intRandomIntegerSix And intRandomIntegerEleven <> intRandomIntegerSeven _
And intRandomIntegerEleven <> intRandomIntegerEight _
And intRandomIntegerEleven <> intRandomIntegerNine And intRandomIntegerEleven <> intRandomIntegerTen _
And intRandomIntegerEleven <> intRandomIntegerTwelve _
And intRandomIntegerEleven <> intRandomIntegerThirteen And intRandomIntegerEleven <> intRandomIntegerFourteen _
And intRandomIntegerEleven <> intRandomIntegerFifteen _
And intRandomIntegerTwelve <> intRandomIntegerOne And intRandomIntegerTwelve <> intRandomIntegerTwo _
And intRandomIntegerTwelve <> intRandomIntegerThree _
And intRandomIntegerTwelve <> intRandomIntegerFour And intRandomIntegerTwelve <> intRandomIntegerFive _
And intRandomIntegerTwelve <> intRandomIntegerSix And intRandomIntegerTwelve <> intRandomIntegerSeven _
And intRandomIntegerTwelve <> intRandomIntegerEight _
And intRandomIntegerTwelve <> intRandomIntegerNine And intRandomIntegerTwelve <> intRandomIntegerTen _
And intRandomIntegerTwelve <> intRandomIntegerEleven _
And intRandomIntegerTwelve <> intRandomIntegerThirteen And intRandomIntegerTwelve <> intRandomIntegerFourteen _
And intRandomIntegerTwelve <> intRandomIntegerFifteen _
And intRandomIntegerThirteen <> intRandomIntegerOne And intRandomIntegerThirteen <> intRandomIntegerTwo _
And intRandomIntegerThirteen <> intRandomIntegerThree _
And intRandomIntegerThirteen <> intRandomIntegerFour And intRandomIntegerThirteen <> intRandomIntegerFive _
And intRandomIntegerThirteen <> intRandomIntegerSix And intRandomIntegerThirteen <> intRandomIntegerSeven _
And intRandomIntegerThirteen <> intRandomIntegerEight _
And intRandomIntegerThirteen <> intRandomIntegerNine And intRandomIntegerThirteen <> intRandomIntegerTen _
And intRandomIntegerThirteen <> intRandomIntegerEleven And intRandomIntegerThirteen <> intRandomIntegerTwelve _
And intRandomIntegerThirteen <> intRandomIntegerFourteen _
And intRandomIntegerThirteen <> intRandomIntegerFifteen _
And intRandomIntegerFourteen <> intRandomIntegerOne And intRandomIntegerFourteen <> intRandomIntegerTwo _
And intRandomIntegerFourteen <> intRandomIntegerThree _
And intRandomIntegerFourteen <> intRandomIntegerFour And intRandomIntegerFourteen <> intRandomIntegerFive _
And intRandomIntegerFourteen <> intRandomIntegerSix And intRandomIntegerFourteen <> intRandomIntegerSeven _
And intRandomIntegerFourteen <> intRandomIntegerEight _
And intRandomIntegerFourteen <> intRandomIntegerNine And intRandomIntegerFourteen <> intRandomIntegerTen _
And intRandomIntegerFourteen <> intRandomIntegerEleven And intRandomIntegerFourteen <> intRandomIntegerTwelve _
And intRandomIntegerFourteen <> intRandomIntegerThirteen _
And intRandomIntegerFourteen <> intRandomIntegerFifteen _
And intRandomIntegerFifteen <> intRandomIntegerOne And intRandomIntegerFifteen <> intRandomIntegerTwo _
And intRandomIntegerFifteen <> intRandomIntegerThree _
And intRandomIntegerFifteen <> intRandomIntegerFour And intRandomIntegerFifteen <> intRandomIntegerFive _
And intRandomIntegerFifteen <> intRandomIntegerSix And intRandomIntegerFifteen <> intRandomIntegerSeven _
And intRandomIntegerFifteen <> intRandomIntegerEight _
And intRandomIntegerFifteen <> intRandomIntegerNine And intRandomIntegerFifteen <> intRandomIntegerTen _
And intRandomIntegerFifteen <> intRandomIntegerEleven And intRandomIntegerFifteen <> intRandomIntegerTwelve _
And intRandomIntegerFifteen <> intRandomIntegerThirteen And intRandomIntegerFifteen <> intRandomIntegerFourteen

        stcQuestionsInUse(0).strQuestion = stcQuestionsDataBase(intRandomIntegerOne).strQuestion
        stcQuestionsInUse(0).strAnswer = stcQuestionsDataBase(intRandomIntegerOne).strCorrectAnswer
        stcQuestionsInUse(0).strBadAnswerOne = stcQuestionDataBase(intRandomIntegerOne).strWrongAnswerOne
        stcQuestionsInUse(0).strBadAnswerTwo = stcQuestionDataBase(intRandomIntegerOne).strWrongAnswerTwo
        stcQuestionsInUse(0).strBadAnswerThree = stcQuestionDataBase(intRandomIntegerOne).strWrongAnswerThree
        stcQuestionsInUse(0).strHintPicture = stcQuestionDataBase(intRandomIntegerOne).strHintPicture
        stcQuestionsInUse(1).strQuestion = stcQuestionsDataBase(intRandomIntegerTwo).strQuestion
        stcQuestionsInUse(1).strAnswer = stcQuestionsDataBase(intRandomIntegerTwo).strCorrectAnswer
        stcQuestionsInUse(1).strBadAnswerOne = stcQuestionDataBase(intRandomIntegerTwo).strWrongAnswerOne
        stcQuestionsInUse(1).strBadAnswerTwo = stcQuestionDataBase(intRandomIntegerTwo).strWrongAnswerTwo
        stcQuestionsInUse(1).strBadAnswerThree = stcQuestionDataBase(intRandomIntegerTwo).strWrongAnswerThree
        stcQuestionsInUse(1).strHintPicture = stcQuestionDataBase(intRandomIntegerTwo).strHintPicture
        stcQuestionsInUse(2).strQuestion = stcQuestionsDataBase(intRandomIntegerThree).strQuestion
        stcQuestionsInUse(2).strAnswer = stcQuestionsDataBase(intRandomIntegerThree).strCorrectAnswer
        stcQuestionsInUse(2).strBadAnswerOne = stcQuestionDataBase(intRandomIntegerThree).strWrongAnswerOne
        stcQuestionsInUse(2).strBadAnswerTwo = stcQuestionDataBase(intRandomIntegerThree).strWrongAnswerTwo
        stcQuestionsInUse(2).strBadAnswerThree = stcQuestionDataBase(intRandomIntegerThree).strWrongAnswerThree
        stcQuestionsInUse(2).strHintPicture = stcQuestionDataBase(intRandomIntegerThree).strHintPicture
        stcQuestionsInUse(3).strQuestion = stcQuestionsDataBase(intRandomIntegerFour).strQuestion
        stcQuestionsInUse(3).strAnswer = stcQuestionsDataBase(intRandomIntegerFour).strCorrectAnswer
        stcQuestionsInUse(3).strBadAnswerOne = stcQuestionDataBase(intRandomIntegerFour).strWrongAnswerOne
        stcQuestionsInUse(3).strBadAnswerTwo = stcQuestionDataBase(intRandomIntegerFour).strWrongAnswerTwo
        stcQuestionsInUse(3).strBadAnswerThree = stcQuestionDataBase(intRandomIntegerFour).strWrongAnswerThree
        stcQuestionsInUse(3).strHintPicture = stcQuestionDataBase(intRandomIntegerFour).strHintPicture
        stcQuestionsInUse(4).strQuestion = stcQuestionsDataBase(intRandomIntegerFive).strQuestion
        stcQuestionsInUse(4).strAnswer = stcQuestionsDataBase(intRandomIntegerFive).strCorrectAnswer
        stcQuestionsInUse(4).strBadAnswerOne = stcQuestionDataBase(intRandomIntegerFive).strWrongAnswerOne
        stcQuestionsInUse(4).strBadAnswerTwo = stcQuestionDataBase(intRandomIntegerFive).strWrongAnswerTwo
        stcQuestionsInUse(4).strBadAnswerThree = stcQuestionDataBase(intRandomIntegerFive).strWrongAnswerThree
        stcQuestionsInUse(4).strHintPicture = stcQuestionDataBase(intRandomIntegerFive).strHintPicture
        stcQuestionsInUse(5).strQuestion = stcQuestionsDataBase(intRandomIntegerSix).strQuestion
        stcQuestionsInUse(5).strAnswer = stcQuestionsDataBase(intRandomIntegerSix).strCorrectAnswer
        stcQuestionsInUse(5).strBadAnswerOne = stcQuestionDataBase(intRandomIntegerSix).strWrongAnswerOne
        stcQuestionsInUse(5).strBadAnswerTwo = stcQuestionDataBase(intRandomIntegerSix).strWrongAnswerTwo
        stcQuestionsInUse(5).strBadAnswerThree = stcQuestionDataBase(intRandomIntegerSix).strWrongAnswerThree
        stcQuestionsInUse(5).strHintPicture = stcQuestionDataBase(intRandomIntegerSix).strHintPicture
        stcQuestionsInUse(6).strQuestion = stcQuestionsDataBase(intRandomIntegerSeven).strQuestion
        stcQuestionsInUse(6).strAnswer = stcQuestionsDataBase(intRandomIntegerSeven).strCorrectAnswer
        stcQuestionsInUse(6).strBadAnswerOne = stcQuestionDataBase(intRandomIntegerSeven).strWrongAnswerOne
        stcQuestionsInUse(6).strBadAnswerTwo = stcQuestionDataBase(intRandomIntegerSeven).strWrongAnswerTwo
        stcQuestionsInUse(6).strBadAnswerThree = stcQuestionDataBase(intRandomIntegerSeven).strWrongAnswerThree
        stcQuestionsInUse(6).strHintPicture = stcQuestionDataBase(intRandomIntegerSeven).strHintPicture
        stcQuestionsInUse(7).strQuestion = stcQuestionsDataBase(intRandomIntegerEight).strQuestion
        stcQuestionsInUse(7).strAnswer = stcQuestionsDataBase(intRandomIntegerEight).strCorrectAnswer
        stcQuestionsInUse(7).strBadAnswerOne = stcQuestionDataBase(intRandomIntegerEight).strWrongAnswerOne
        stcQuestionsInUse(7).strBadAnswerTwo = stcQuestionDataBase(intRandomIntegerEight).strWrongAnswerTwo
        stcQuestionsInUse(7).strBadAnswerThree = stcQuestionDataBase(intRandomIntegerEight).strWrongAnswerThree
        stcQuestionsInUse(7).strHintPicture = stcQuestionDataBase(intRandomIntegerEight).strHintPicture
        stcQuestionsInUse(8).strQuestion = stcQuestionsDataBase(intRandomIntegerNine).strQuestion
        stcQuestionsInUse(8).strAnswer = stcQuestionsDataBase(intRandomIntegerNine).strCorrectAnswer
        stcQuestionsInUse(8).strBadAnswerOne = stcQuestionDataBase(intRandomIntegerNine).strWrongAnswerOne
        stcQuestionsInUse(8).strBadAnswerTwo = stcQuestionDataBase(intRandomIntegerNine).strWrongAnswerTwo
        stcQuestionsInUse(8).strBadAnswerThree = stcQuestionDataBase(intRandomIntegerNine).strWrongAnswerThree
        stcQuestionsInUse(8).strHintPicture = stcQuestionDataBase(intRandomIntegerNine).strHintPicture
        stcQuestionsInUse(9).strQuestion = stcQuestionsDataBase(intRandomIntegerTen).strQuestion
        stcQuestionsInUse(9).strAnswer = stcQuestionsDataBase(intRandomIntegerTen).strCorrectAnswer
        stcQuestionsInUse(9).strBadAnswerOne = stcQuestionDataBase(intRandomIntegerTen).strWrongAnswerOne
        stcQuestionsInUse(9).strBadAnswerTwo = stcQuestionDataBase(intRandomIntegerTen).strWrongAnswerTwo
        stcQuestionsInUse(9).strBadAnswerThree = stcQuestionDataBase(intRandomIntegerTen).strWrongAnswerThree
        stcQuestionsInUse(9).strHintPicture = stcQuestionDataBase(intRandomIntegerTen).strHintPicture
        stcQuestionsInUse(10).strQuestion = stcQuestionsDataBase(intRandomIntegerEleven).strQuestion
        stcQuestionsInUse(10).strAnswer = stcQuestionsDataBase(intRandomIntegerEleven).strCorrectAnswer
        stcQuestionsInUse(10).strBadAnswerOne = stcQuestionDataBase(intRandomIntegerEleven).strWrongAnswerOne
        stcQuestionsInUse(10).strBadAnswerTwo = stcQuestionDataBase(intRandomIntegerEleven).strWrongAnswerTwo
        stcQuestionsInUse(10).strBadAnswerThree = stcQuestionDataBase(intRandomIntegerEleven).strWrongAnswerThree
        stcQuestionsInUse(10).strHintPicture = stcQuestionDataBase(intRandomIntegerEleven).strHintPicture
        stcQuestionsInUse(11).strQuestion = stcQuestionsDataBase(intRandomIntegerTwelve).strQuestion
        stcQuestionsInUse(11).strAnswer = stcQuestionsDataBase(intRandomIntegerTwelve).strCorrectAnswer
        stcQuestionsInUse(11).strBadAnswerOne = stcQuestionDataBase(intRandomIntegerTwelve).strWrongAnswerOne
        stcQuestionsInUse(11).strBadAnswerTwo = stcQuestionDataBase(intRandomIntegerTwelve).strWrongAnswerTwo
        stcQuestionsInUse(11).strBadAnswerThree = stcQuestionDataBase(intRandomIntegerTwelve).strWrongAnswerThree
        stcQuestionsInUse(11).strHintPicture = stcQuestionDataBase(intRandomIntegerTwelve).strHintPicture
        stcQuestionsInUse(12).strQuestion = stcQuestionsDataBase(intRandomIntegerThirteen).strQuestion
        stcQuestionsInUse(12).strAnswer = stcQuestionsDataBase(intRandomIntegerThirteen).strCorrectAnswer
        stcQuestionsInUse(12).strBadAnswerOne = stcQuestionDataBase(intRandomIntegerThirteen).strWrongAnswerOne
        stcQuestionsInUse(12).strBadAnswerTwo = stcQuestionDataBase(intRandomIntegerThirteen).strWrongAnswerTwo
        stcQuestionsInUse(12).strBadAnswerThree = stcQuestionDataBase(intRandomIntegerThirteen).strWrongAnswerThree
        stcQuestionsInUse(12).strHintPicture = stcQuestionDataBase(intRandomIntegerThirteen).strHintPicture
        stcQuestionsInUse(13).strQuestion = stcQuestionsDataBase(intRandomIntegerFourteen).strQuestion
        stcQuestionsInUse(13).strAnswer = stcQuestionsDataBase(intRandomIntegerFourteen).strCorrectAnswer
        stcQuestionsInUse(13).strBadAnswerOne = stcQuestionDataBase(intRandomIntegerFourteen).strWrongAnswerOne
        stcQuestionsInUse(13).strBadAnswerTwo = stcQuestionDataBase(intRandomIntegerFourteen).strWrongAnswerTwo
        stcQuestionsInUse(13).strBadAnswerThree = stcQuestionDataBase(intRandomIntegerFourteen).strWrongAnswerThree
        stcQuestionsInUse(13).strHintPicture = stcQuestionDataBase(intRandomIntegerFourteen).strHintPicture
        stcQuestionsInUse(14).strQuestion = stcQuestionsDataBase(intRandomIntegerFifteen).strQuestion
        stcQuestionsInUse(14).strAnswer = stcQuestionsDataBase(intRandomIntegerFifteen).strCorrectAnswer
        stcQuestionsInUse(14).strBadAnswerOne = stcQuestionDataBase(intRandomIntegerFifteen).strWrongAnswerOne
        stcQuestionsInUse(14).strBadAnswerTwo = stcQuestionDataBase(intRandomIntegerFifteen).strWrongAnswerTwo
        stcQuestionsInUse(14).strBadAnswerThree = stcQuestionDataBase(intRandomIntegerFifteen).strWrongAnswerThree
        stcQuestionsInUse(14).strHintPicture = stcQuestionDataBase(intRandomIntegerFifteen).strHintPicture

    End Sub

    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    'Checks the answer clicked by the user if correct then calls load next question and increases intMoney won based on 
    'the select case. Changes the color of the button depending if the user clicked the correct answer (green) or the wrong 
    'one (red)
    '
    'post: If the blnIsWrong = false then increasses value of intMoneyWon based on the select case, if bln is Wrong = true
    'then shows a message box and calls resetGame to reset the game. And if the intCurrentQuestion = 14 and blnIsWrong = false
    'endGame timer is enabled and askforuserrespose is called  
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Sub CheckAnswer(ByVal strAnswerChosen As String, ByVal blnIsWrong As Boolean, _
                    ByRef stcQuestions() As QuestionsInUse, ByRef btnOptions() As Button, _
                    ByRef intCurrentQuestion As Integer, ByVal chrButtonClickedTag As Char, _
                    ByRef lblQuestionBeingAsked As Label, ByRef lstWonMoney As ListBox, _
                    ByRef lblCountDownTimer As Label, ByRef tmrTimer As Timer, _
                    ByRef lblLabels() As Label, ByRef btnButtons() As Button, ByRef mnuAwayWalk As ToolStripMenuItem,
                    ByRef mnuLineLife As ToolStripMenuItem, ByRef picHint As PictureBox, ByRef prgLeft As ProgressBar, _
                    ByRef prgRight As ProgressBar, ByRef tmrEndGame As Timer)

        Dim strRightAnswer As String
        'Dim intLoop As Integer
        Dim strUserReponse As String
        'Dim intRandomIneger As Integer

        strRightAnswer = stcQuestions(intCurrentQuestion).strAnswer 'Compares the answer chosen by the player to the correct answer in the array
        blnIsWrong = CorrectAnswer(strAnswerChosen, strRightAnswer)
        'strAnswer = MsgBox("Congratulations You have won " & intMAXMONEY & "Would you like to play again", vbYesNo)

        If blnIsWrong = False And chrButtonClickedTag = "a" Then
            btnOptions(0).BackColor = Color.Lime
        ElseIf blnIsWrong = False And chrButtonClickedTag = "b" Then
            btnOptions(1).BackColor = Color.Lime
        ElseIf blnIsWrong = False And chrButtonClickedTag = "c" Then
            btnOptions(2).BackColor = Color.Lime
        ElseIf blnIsWrong = False And chrButtonClickedTag = "d" Then
            btnOptions(3).BackColor = Color.Lime
        ElseIf blnIsWrong = True And chrButtonClickedTag = "a" Then
            btnOptions(0).BackColor = Color.Red
        ElseIf blnIsWrong = True And chrButtonClickedTag = "b" Then
            btnOptions(1).BackColor = Color.Red
        ElseIf blnIsWrong = True And chrButtonClickedTag = "c" Then
            btnOptions(2).BackColor = Color.Red
        ElseIf blnIsWrong = True And chrButtonClickedTag = "d" Then
            btnOptions(3).BackColor = Color.Red
        End If

        If blnIsWrong = False And intMoneyWon <> intMAXMONEY And intCurrentQuestion <> 14 Then
            Select Case intCurrentQuestion
                Case 0
                    intMoneyWon = 100
                Case 1
                    intMoneyWon = 500
                Case 2
                    intMoneyWon = 1000
                Case 3
                    intMoneyWon = 2500
                Case 4
                    intMoneyWon = 5000
                Case 5
                    intMoneyWon = 10000
                Case 6
                    intMoneyWon = 20000
                Case 7
                    intMoneyWon = 40000
                Case 8
                    intMoneyWon = 80000
                Case 9
                    intMoneyWon = 100000
                Case 10
                    intMoneyWon = 500000
                Case 11
                    intMoneyWon = 1000000
                Case 12
                    intMoneyWon = 5000000
                Case 13
                    intMoneyWon = 10000000
                Case 14
                    intMoneyWon = 15000000
            End Select
            MessageBox.Show("Correct!, You have won " & Format(intMoneyWon, "Currency"))
            Call SelectItem(lstWonMoney, intCurrentQuestion)
            intCurrentQuestion = intCurrentQuestion + 1
            Call LoadNextQuestion(stcQuestionsInUse, intCurrentQuestion, btnOptions, _
                                  lblQuestionBeingAsked, intMoneyWon, lblCountDownTimer, tmrTimer, picHint)
        ElseIf blnIsWrong = True Then
            MessageBox.Show("Wrong!, You have lost " & Format(intMoneyWon, "Currency") & " Game Over!")
            Call ResetGame(btnButtons, lblLabels, tmrTimer, picHint, lstWonMoney, _
            mnuLineLife, mnuAwayWalk, tmrEndGame, stcQuestions, prgLeft, prgRight)
            'lblAllLabels(0).Visible = False 'lbloptiona
            'lblAllLabels(1).Visible = False 'lbloptionb
            'lblAllLabels(2).Visible = False 'lbloptionc
            'lblAllLabels(3).Visible = False 'lbloptiond
            'lblAllLabels(4).Visible = False 'lbltimer
            'lblAllLabels(5).Visible = False 'lblquestion
            'btnAllButtons(0).Visible = False 'btnanswer a
            'btnAllButtons(1).Visible = False 'btnanswerb
            'btnAllButtons(2).Visible = False 'btnanswerc
            'btnAllButtons(3).Visible = False 'btnanswerd
            'btnAllButtons(4).Visible = False 'btnLifeline
            'btnAllButtons(4).Enabled = False 'btnlifeline
            'btnAllButtons(5).Visible = False 'btnwalkaway
            'btnAllButtons(6).Visible = True 'btnStartNewGame
            'mnuLineLife.Visible = False
            'mnuAwayWalk.Visible = False
            'picHint.Visible = False
            'lstWonMoney.Visible = False
            'lstWonMoney.SelectedIndex = 0
            'tmrTimer.Enabled = False
            'tmrEndGame.Enabled = False
            'prgLeft.Visible = False
            'prgRight.Visible = False
            'For intLoop = 0 To stcQuestions.Length - 1
            '    stcQuestions(intLoop).strAnswer = Nothing
            '    stcQuestions(intLoop).strQuestion = Nothing
            '    stcQuestions(intLoop).strHintPicture = Nothing
            '    stcQuestions(intLoop).strBadAnswerOne = Nothing
            '    stcQuestions(intLoop).strBadAnswerTwo = Nothing
            '    stcQuestions(intLoop).strBadAnswerThree = Nothing
            'Next intLoop
        ElseIf blnIsWrong = False And intCurrentQuestion = 14 Then
            tmrEndGame.Enabled = True
            Call SelectItem(lstWonMoney, intCurrentQuestion)
            Call AskForUserResponse("Correct!, Congratulations! You have won " & Format(intMAXMONEY, "Currency") _
            & " Would you like to play again?", strUserReponse)
            'strUserReponse = MsgBox("Correct!, Congratulations! You have won " & Format(intMAXMONEY, "Currency") _
            '       & " Would you like to play again?", +vbYesNo)
            If strUserReponse = vbYes Then
                Call ResetGame(btnButtons, lblLabels, tmrTimer, picHint, lstWonMoney, _
            mnuLineLife, mnuAwayWalk, tmrEndGame, stcQuestions, prgLeft, prgRight)
            Else
                End
            End If
        End If

    End Sub

    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    'Displays the next question and shows the hint picture in que from the stcQuestionsInUse array using the 
    'integer provided by intQurrentQuestion as index value.
    '
    'post: Displays the question next in que using intCurrentQuestion as the index value, also displays the hint 
    'picture if any int the element.
    'calls randomize answers to randomize the order the correct answer and the bad answers are shown on the buttons
    'clears the back color of the buttons. Enables the tmrQuestionTimer
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Sub LoadNextQuestion(ByRef stcQuestionsInUse() As QuestionsInUse, ByRef intCurrentQuestion As Integer, _
                         ByRef btnOptions() As Button, ByRef lblQuestionBeingAsked As Label, _
                         ByVal intMoneyWon As Integer, ByRef lblCountDownTimer As Label, ByRef tmrTimer As Timer, _
                         ByRef picHint As PictureBox)

        'If intCurrentQuestion <> 14 Then
        Call RandomizeAnswers(stcQuestionsInUse, btnOptions, intCurrentQuestion)
        intTimeLeft = intTIMEONQUESTION
        picHint.Image = Nothing
        btnOptions(0).BackColor = Nothing
        btnOptions(1).BackColor = Nothing
        btnOptions(2).BackColor = Nothing
        btnOptions(3).BackColor = Nothing
        tmrTimer.Enabled = True

        lblQuestionBeingAsked.Text = stcQuestionsInUse(intCurrentQuestion).strQuestion

        If stcQuestionsInUse(intCurrentQuestion).strHintPicture <> Nothing Then
            picHint.Image = Image.FromFile(stcQuestionsInUse(intCurrentQuestion).strHintPicture)
        Else
            picHint.Image = Image.FromFile("WhoWantsToBeAMillionare.png")
        End If
        'Else
        '    Call RandomizeAnswers(stcQuestionsInUse, btnOptions, 14)
        '    intTimeLeft = 60
        '    picHint.Image = Nothing
        '    btnOptions(0).BackColor = Nothing
        '    btnOptions(1).BackColor = Nothing
        '    btnOptions(2).BackColor = Nothing
        '    btnOptions(3).BackColor = Nothing
        '    tmrTimer.Enabled = True
        '    lblQuestionBeingAsked.Text = stcQuestionsInUse(intCurrentQuestion).strQuestion
        '    If stcQuestionsInUse(intCurrentQuestion).strHintPicture <> Nothing Then
        '        picHint.Image = Image.FromFile(stcQuestionsInUse(intCurrentQuestion).strHintPicture)
        '    End If
        'End If
        'Do
        '    intRandomInteger = RndInt(3, 1)
        '    If intRandomInteger = 0 Then
        '        btnOptions(0).Text = stcQuestionsInUse(intCurrentQuestion).strAnswer
        '    ElseIf intRandomInteger = 1 Then
        '        btnOptions(1).Text = stcQuestionsInUse(intCurrentQuestion).strAnswer
        '    ElseIf intRandomInteger = 2 Then
        '        btnOptions(2).Text = stcQuestionsInUse(intCurrentQuestion).strAnswer
        '    ElseIf intRandomInteger = 3 Then
        '        btnOptions(3).Text = stcQuestionsInUse(intCurrentQuestion).strAnswer
        '    End If
        'Loop While btnOptions(0).Text = btnOptions(1).Text Or btnOptions(0).Text = btnOptions(2).Text Or btnOptions(0).Text = btnOptions(3).Text _
        '    Or btnOptions(1).Text = btnOptions(2).Text Or btnOptions(1).Text = btnOptions(3).Text Or btnOptions(2).Text = btnOptions(3).Text
    End Sub

    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    'Returns the a boolean value (true) if strAnswerChosen is equal to strCorrectAnswer 
    '
    'Post: A boolean is returned
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Function CorrectAnswer(ByVal strAnswerChosen As String, ByVal strCorrectAnswer As String) As Boolean

        If strAnswerChosen = strCorrectAnswer Then
            Return False
        Else
            Return True
        End If

    End Function

    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    'Returns a random number from the range provided by intHighNum and intLowNum 
    '
    'post: A random number has been returned.
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Function RndInt(ByVal intHighNum As Integer, ByVal intLowNum As Integer) As Integer

        Randomize()
        Return (Int((intHighNum - intLowNum + 1) * Rnd() + intLowNum))

    End Function

    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    'Randomizes which answer is shown an a button 
    '
    'post: The order the asnwers are shown on the four buttons is randomized
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Sub RandomizeAnswers(ByRef stcQuestionsInUse() As QuestionsInUse, _
                         ByRef btnOptions() As Button, ByVal intCurrentQuestion As Integer)

        Dim intRandomIntegerOne As Integer
        Dim intRandomIntegerTwo As Integer
        Dim intRandomIntegerThree As Integer
        Dim intRandomIntegerFour As Integer

        btnOptions(0).Text = Nothing
        btnOptions(1).Text = Nothing
        btnOptions(2).Text = Nothing
        btnOptions(3).Text = Nothing
        intRandomIntegerOne = Nothing
        intRandomIntegerTwo = Nothing
        intRandomIntegerThree = Nothing
        intRandomIntegerFour = Nothing

        Do
            intRandomIntegerOne = RndInt(3, 0)
            If intRandomIntegerOne = 0 Then
                btnOptions(0).Text = stcQuestionsInUse(intCurrentQuestion).strAnswer
            ElseIf intRandomIntegerOne = 1 Then
                btnOptions(1).Text = stcQuestionsInUse(intCurrentQuestion).strAnswer
            ElseIf intRandomIntegerOne = 2 Then
                btnOptions(2).Text = stcQuestionsInUse(intCurrentQuestion).strAnswer
            ElseIf intRandomIntegerOne = 3 Then
                btnOptions(3).Text = stcQuestionsInUse(intCurrentQuestion).strAnswer
            End If
            intRandomIntegerTwo = RndInt(3, 0)
            If intRandomIntegerTwo = 0 Then
                btnOptions(0).Text = stcQuestionsInUse(intCurrentQuestion).strBadAnswerOne
            ElseIf intRandomIntegerTwo = 1 Then
                btnOptions(1).Text = stcQuestionsInUse(intCurrentQuestion).strBadAnswerOne
            ElseIf intRandomIntegerTwo = 2 Then
                btnOptions(2).Text = stcQuestionsInUse(intCurrentQuestion).strBadAnswerOne
            ElseIf intRandomIntegerTwo = 3 Then
                btnOptions(3).Text = stcQuestionsInUse(intCurrentQuestion).strBadAnswerOne
            End If
            intRandomIntegerThree = RndInt(3, 0)
            If intRandomIntegerThree = 0 Then
                btnOptions(0).Text = stcQuestionsInUse(intCurrentQuestion).strBadAnswerTwo
            ElseIf intRandomIntegerThree = 1 Then
                btnOptions(1).Text = stcQuestionsInUse(intCurrentQuestion).strBadAnswerTwo
            ElseIf intRandomIntegerThree = 2 Then
                btnOptions(2).Text = stcQuestionsInUse(intCurrentQuestion).strBadAnswerTwo
            ElseIf intRandomIntegerThree = 3 Then
                btnOptions(3).Text = stcQuestionsInUse(intCurrentQuestion).strBadAnswerTwo
            End If
            intRandomIntegerFour = RndInt(3, 0)
            If intRandomIntegerFour = 0 Then
                btnOptions(0).Text = stcQuestionsInUse(intCurrentQuestion).strBadAnswerThree
            ElseIf intRandomIntegerFour = 1 Then
                btnOptions(1).Text = stcQuestionsInUse(intCurrentQuestion).strBadAnswerThree
            ElseIf intRandomIntegerFour = 2 Then
                btnOptions(2).Text = stcQuestionsInUse(intCurrentQuestion).strBadAnswerThree
            ElseIf intRandomIntegerFour = 3 Then
                btnOptions(3).Text = stcQuestionsInUse(intCurrentQuestion).strBadAnswerThree
            End If
        Loop Until btnOptions(0).Text <> btnOptions(1).Text And btnOptions(0).Text <> btnOptions(2).Text _
            And btnOptions(0).Text <> btnOptions(3).Text And btnOptions(1).Text <> btnOptions(0).Text _
            And btnOptions(1).Text <> btnOptions(2).Text And btnOptions(1).Text <> btnOptions(3).Text _
            And btnOptions(2).Text <> btnOptions(1).Text And btnOptions(2).Text <> btnOptions(3).Text _
            And btnOptions(2).Text <> btnOptions(0).Text And btnOptions(3).Text <> btnOptions(0).Text _
            And btnOptions(3).Text <> btnOptions(1).Text And btnOptions(3).Text <> btnOptions(2).Text _
            And btnOptions(0).Text <> Nothing And btnOptions(1).Text <> Nothing And btnOptions(2).Text <> Nothing _
            And btnOptions(3).Text <> Nothing

    End Sub

    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    'Selects the Current question the player is on, on the lstMoneyWon listbox
    '
    'post: The integer provided by intCurrentQuestion + 1 is used for the 
    'SelectedIndex value for lstMoneyWon list box 
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Sub SelectItem(ByRef lstWonMoney As ListBox, ByVal intCurrentQuestion As Integer)

        lstWonMoney.SelectedIndex = intCurrentQuestion + 1

    End Sub

    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    'When clicked 2 of the wrong answers shown on the 4 buttons are colored red  
    '
    'post: 2 random integers are created and then used for the index value for the btnOptions array then checked 
    'in the do loop until not one of the text on the buttons is equal to the strCorrectAnswer. mnuLifeLine and 
    'btnLifeLine are disabled
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Private Sub btnLifeLine_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnLifeLine.Click

        Dim btnOptionsThree() As Button = {Me.btnAnswerChoiceA, Me.btnAnswerChoiceB, Me.btnAnswerChoiceC, Me.btnAnswerChoiceD}
        Dim intRandomIntegerOne As Integer
        Dim intRandomIntegerTwo As Integer

        Do
            intRandomIntegerOne = RndInt(3, 0)
            intRandomIntegerTwo = RndInt(3, 0)
        Loop Until intRandomIntegerOne <> intRandomIntegerTwo And intRandomIntegerTwo <> intRandomIntegerOne _
            And btnOptionsThree(intRandomIntegerOne).Text <> stcQuestionsInUse(intQuestionPlayerIsOn).strAnswer _
            And btnOptionsThree(intRandomIntegerTwo).Text <> stcQuestionsInUse(intQuestionPlayerIsOn).strAnswer

        If intRandomIntegerOne = 0 Then
            btnOptionsThree(0).BackColor = Color.Red
        ElseIf intRandomIntegerOne = 1 Then
            btnOptionsThree(1).BackColor = Color.Red
        ElseIf intRandomIntegerOne = 2 Then
            btnOptionsThree(2).BackColor = Color.Red
        ElseIf intRandomIntegerOne = 3 Then
            btnOptionsThree(3).BackColor = Color.Red
        End If
        If intRandomIntegerTwo = 0 Then
            btnOptionsThree(0).BackColor = Color.Red
        ElseIf intRandomIntegerTwo = 1 Then
            btnOptionsThree(1).BackColor = Color.Red
        ElseIf intRandomIntegerTwo = 2 Then
            btnOptionsThree(2).BackColor = Color.Red
        ElseIf intRandomIntegerTwo = 3 Then
            btnOptionsThree(3).BackColor = Color.Red
        End If

        Me.btnLifeLine.Enabled = False
        Me.mnuLifeLine.Enabled = False

    End Sub

    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    'Checks if the user has more than or equal to 20000 in the integer intMoneyWon if the player dose then a message box is shown 
    'to ask if the user is sure about walking away then if yes then resetGame is called to reset the game
    '
    'post: the Player is asked if they are sure if they want to walk away with their current amount of money, if yes then the game 
    'is reset, else the game resumes.
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Private Sub btnWalkAway_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnWalkAway.Click, mnuWalkAway.Click

        Dim strUserResponse As String
        Dim lblAllLabelsOnForm() As Label = {Me.lblOptionA, Me.lblOptionB, Me.lblOptionC, Me.lblOptionD, Me.lblTimer, Me.lblQuestion}
        Dim btnAllButtonsOnForm() As Button = {Me.btnAnswerChoiceA, Me.btnAnswerChoiceB, Me.btnAnswerChoiceC, Me.btnAnswerChoiceD, _
                                               Me.btnLifeLine, Me.btnWalkAway, Me.btnStartNewGame}

        If intMoneyWon > 20000 Then
            'strUserResponse = MsgBox("Are you sure you want to walk away with " & Format(intMoneyWon, "Currency"), vbYesNo)
            Call AskForUserResponse("Are you sure you want to walk away with " & Format(intMoneyWon, "Currency"), strUserResponse)
            If strUserResponse = vbYes Then
                MessageBox.Show("Congratulations You have won " & Format(intMoneyWon, "Currency") & " Have Fun!")
                intMoneyWon = 0
                Call ResetGame(btnAllButtonsOnForm, lblAllLabelsOnForm, Me.tmrQuestionTimer, Me.picHintPicture, Me.lstMoneyWon, _
            Me.mnuLifeLine, Me.mnuWalkAway, Me.tmrEndGameCelebration, stcQuestionsInUse, Me.prgTimeLeft, Me.prgTimeRight)
            Else
                MessageBox.Show("Keep Playing, Good Luck")
            End If
        Else
            MessageBox.Show("Sorry You Only Have " & Format(intMoneyWon, "Currency") & " You Need At Least $20000 To Walk Away")
        End If

    End Sub

    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    'When the Player runs out of time a message box is shown and AskForUserResponse is called to ask if the user would 
    'like to play again, and if yes ResetGame is called.
    '
    'post: When lblTimer or prgTimeLeft Or prgTimeRight is 0 a message box is shown and AskForUserResponse is called 
    'to ask if the user would like to play again, and if yes ResetGame is called.
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Private Sub tmrQuestionTimer_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles tmrQuestionTimer.Tick

        Dim decProgressBarDivisor As Decimal    'The number the progress bar value will be decreased by
        Dim strUserReponse As String
        Dim lblAllLabelsOnForm() As Label = {Me.lblOptionA, Me.lblOptionB, Me.lblOptionC, Me.lblOptionD, Me.lblTimer, Me.lblQuestion}
        Dim btnAllButtonsOnForm() As Button = {Me.btnAnswerChoiceA, Me.btnAnswerChoiceB, Me.btnAnswerChoiceC, Me.btnAnswerChoiceD, _
                                               Me.btnLifeLine, Me.btnWalkAway, Me.btnStartNewGame}

        decProgressBarDivisor = 100 / intTIMEONQUESTION
        intTimeLeft = intTimeLeft - 1

        Me.lblTimer.Text = intTimeLeft
        Me.prgTimeRight.Value = intTimeLeft * decProgressBarDivisor 'I divided 100 (the full value of the progres bar) by 60 (the Seconds in one Minute) 
        Me.prgTimeLeft.Value = intTimeLeft * decProgressBarDivisor  'To Illustrate time in the proggress bar

        If Me.lblTimer.Text = 0 Or Me.prgTimeLeft.Value = 0 Or Me.prgTimeRight.Value = 0 Then
            Me.tmrQuestionTimer.Enabled = False
            MessageBox.Show("Bad Luck, You Ran Out Of Time. Good Luck Next Time!")
            Call AskForUserResponse(" Would you like to play again?", strUserReponse)
            'strUserReponse = MsgBox("Correct!, Congratulations! You have won " & Format(intMAXMONEY, "Currency") _
            '& " Would you like to play again?", +vbYesNo)
            If strUserReponse = vbYes Then
                Call ResetGame(btnAllButtonsOnForm, lblAllLabelsOnForm, Me.tmrQuestionTimer, Me.picHintPicture, Me.lstMoneyWon, _
            Me.mnuLifeLine, Me.mnuWalkAway, Me.tmrEndGameCelebration, stcQuestionsInUse, Me.prgTimeLeft, Me.prgTimeRight)
                'lblAllLabels(0).Visible = False 'lbloptiona
                'lblAllLabels(1).Visible = False 'lbloptionb
                'lblAllLabels(2).Visible = False 'lbloptionc
                'lblAllLabels(3).Visible = False 'lbloptiond
                'lblAllLabels(4).Visible = False 'lbltimer
                'lblAllLabels(5).Visible = False 'lblquestion
                'btnAllButtons(0).Visible = False 'btnanswer a
                'btnAllButtons(1).Visible = False 'btnanswerb
                'btnAllButtons(2).Visible = False 'btnanswerc
                'btnAllButtons(3).Visible = False 'btnanswerd
                'btnAllButtons(4).Visible = False 'btnLifeline
                'btnAllButtons(4).Enabled = False 'btnlifeline
                'btnAllButtons(5).Visible = False 'btnwalkaway
                'btnAllButtons(6).Visible = True 'btnStartNewGame
                'mnuLineLife.Visible = False
                'mnuAwayWalk.Visible = False
                'picHint.Visible = False
                'lstWonMoney.Visible = False
                'lstWonMoney.SelectedIndex = 0
                'tmrTimer.Enabled = False
                'tmrEndGame.Enabled = False
                'prgLeft.Visible = False
                'prgRight.Visible = False
                'For intLoop = 0 To stcQuestions.Length - 1
                '    stcQuestions(intLoop).strAnswer = Nothing
                '    stcQuestions(intLoop).strQuestion = Nothing
                '    stcQuestions(intLoop).strHintPicture = Nothing
                '    stcQuestions(intLoop).strBadAnswerOne = Nothing
                '    stcQuestions(intLoop).strBadAnswerTwo = Nothing
                '    stcQuestions(intLoop).strBadAnswerThree = Nothing
                'Next intLoop
            Else
                MyBase.Close()
            End If
            'Me.lblOptionA.Visible = False
            'Me.lblOptionB.Visible = False
            'Me.lblOptionC.Visible = False
            'Me.lblOptionD.Visible = False
            'Me.btnAnswerChoiceA.Visible = False
            'Me.btnAnswerChoiceB.Visible = False
            'Me.btnAnswerChoiceC.Visible = False
            'Me.btnAnswerChoiceD.Visible = False
            'Me.mnuLifeLine.Visible = False
            'Me.btnLifeLine.Enabled = False
            'Me.mnuWalkAway.Visible = False
            'Me.btnWalkAway.Visible = False
            'Me.btnLifeLine.Visible = False
            'Me.picHintPicture.Visible = False
            'Me.lstMoneyWon.Visible = False
            'Me.lblQuestion.Visible = False
            'Me.lblTimer.Visible = False
            'Me.btnStartNewGame.Visible = True
            'Me.lstMoneyWon.SelectedIndex = 0
            'Me.tmrQuestionTimer.Enabled = False
            'Me.tmrEndGameCelebration.Enabled = False
            'Me.prgTimeLeft.Visible = False
            'Me.prgTimeRight.Visible = False
        End If

    End Sub

    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    'Resets all objects to their default state
    '
    'post: lblAllLabels is visable is all set to false, btnAllbuttons is visable is all set to false,
    'mnuLifeLine and mnuWalkAway is visable is set to false both progress bars are visable is set to false
    'lstMoneyWon is visable is set to false tmrEndGame is disabled tmrTimer is disabled, stcQuestionsInUse is 
    'cleared Using a for next loop
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Sub ResetGame(ByRef btnButtons() As Button, ByRef lblLabels() As Label, ByRef tmrTimer As Timer, _
                  ByRef picHint As PictureBox,  ByRef lstWonMoney As ListBox, _
                   ByRef mnuLineLife As ToolStripMenuItem, ByRef mnuAwayWalk As ToolStripMenuItem, _
                   ByRef tmrEndGame As Timer, ByRef stcQuestions() As QuestionsInUse, ByRef prgLeft As ProgressBar _
                   , ByRef prgRight As ProgressBar)

        lblLabels(0).Visible = False    'lbloption a
        lblLabels(1).Visible = False    'lbloption b
        lblLabels(2).Visible = False    'lbloption c
        lblLabels(3).Visible = False    'lbloption d
        lblLabels(4).Visible = False    'lbltimer
        lblLabels(5).Visible = False    'lblquestion
        btnButtons(0).Visible = False   'btnanswerchoice a
        btnButtons(1).Visible = False   'btnanswerchoice b
        btnButtons(2).Visible = False   'btnanswerchoice c
        btnButtons(3).Visible = False   'btnanswerchoice d
        mnuLineLife.Visible = False     'mnulifeline
        btnButtons(4).Visible = False   'btnlifeline
        btnButtons(4).Enabled = False   'btnlifeline
        mnuAwayWalk.Visible = False     'mnuwalkaway
        btnButtons(5).Visible = False   'btnwalkaway
        btnButtons(5).Visible = False   'btnwalkaway
        picHint.Visible = False         'pichintpicture
        lstWonMoney.Visible = False     'lstmoneywon
        btnButtons(6).Visible = True    'btnStartNewGame
        tmrEndGame.Enabled = False      'tmrendgamecelebration
        tmrTimer.Enabled = False        'tmrQuestionTimer
        prgLeft.Visible = False         'prgTimeLeft
        prgRight.Visible = False        'prgTimeRight
        intMoneyWon = 0                 'intMoneyWon
        For intLoop = 0 To stcQuestions.Length - 1
            stcQuestions(intLoop).strAnswer = Nothing
            stcQuestions(intLoop).strQuestion = Nothing
            stcQuestions(intLoop).strHintPicture = Nothing
            stcQuestions(intLoop).strBadAnswerOne = Nothing
            stcQuestions(intLoop).strBadAnswerTwo = Nothing
            stcQuestions(intLoop).strBadAnswerThree = Nothing
        Next intLoop

    End Sub

    Private Sub tmrEndGameCelebration_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles tmrEndGameCelebration.Tick

        Dim intRandomInteger As Integer

        intRandomInteger = RndInt(10, 1)

        Select Case intRandomInteger
            Case 1
                Me.btnAnswerChoiceA.BackColor = Color.Red
                Me.btnAnswerChoiceB.BackColor = Color.Red
                Me.btnAnswerChoiceC.BackColor = Color.Red
                Me.btnAnswerChoiceD.BackColor = Color.Red
            Case 2
                Me.btnAnswerChoiceA.BackColor = Color.Orange
                Me.btnAnswerChoiceB.BackColor = Color.Orange
                Me.btnAnswerChoiceC.BackColor = Color.Orange
                Me.btnAnswerChoiceD.BackColor = Color.Orange
            Case 3
                Me.btnAnswerChoiceA.BackColor = Color.Yellow
                Me.btnAnswerChoiceB.BackColor = Color.Yellow
                Me.btnAnswerChoiceC.BackColor = Color.Yellow
                Me.btnAnswerChoiceD.BackColor = Color.Yellow
            Case 4
                Me.btnAnswerChoiceA.BackColor = Color.Blue
                Me.btnAnswerChoiceB.BackColor = Color.Blue
                Me.btnAnswerChoiceC.BackColor = Color.Blue
                Me.btnAnswerChoiceD.BackColor = Color.Blue
            Case 5
                Me.btnAnswerChoiceA.BackColor = Color.Purple
                Me.btnAnswerChoiceB.BackColor = Color.Purple
                Me.btnAnswerChoiceC.BackColor = Color.Purple
                Me.btnAnswerChoiceD.BackColor = Color.Purple
            Case 6
                Me.btnAnswerChoiceA.BackColor = Color.Green
                Me.btnAnswerChoiceB.BackColor = Color.Green
                Me.btnAnswerChoiceC.BackColor = Color.Green
                Me.btnAnswerChoiceD.BackColor = Color.Green
            Case 7
                Me.btnAnswerChoiceA.BackColor = Color.Violet
                Me.btnAnswerChoiceB.BackColor = Color.Violet
                Me.btnAnswerChoiceC.BackColor = Color.Violet
                Me.btnAnswerChoiceD.BackColor = Color.Violet
            Case 8
                Me.btnAnswerChoiceA.BackColor = Color.DarkCyan
                Me.btnAnswerChoiceB.BackColor = Color.DarkCyan
                Me.btnAnswerChoiceC.BackColor = Color.DarkCyan
                Me.btnAnswerChoiceD.BackColor = Color.DarkCyan
            Case 9
                Me.btnAnswerChoiceA.BackColor = Color.Lime
                Me.btnAnswerChoiceB.BackColor = Color.Lime
                Me.btnAnswerChoiceC.BackColor = Color.Lime
                Me.btnAnswerChoiceD.BackColor = Color.Lime
            Case 10
                Me.btnAnswerChoiceA.BackColor = Color.Magenta
                Me.btnAnswerChoiceB.BackColor = Color.Magenta
                Me.btnAnswerChoiceC.BackColor = Color.Magenta
                Me.btnAnswerChoiceD.BackColor = Color.Magenta
        End Select

    End Sub

    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    'Shows a message box with the strStatement and a vbYesNo buttons
    '
    'post: strUserAnswer Stores What the User responded with.
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Sub AskForUserResponse(ByVal strStatement As String, ByRef strUserAnswer As String)

        strUserAnswer = MsgBox(strStatement, vbYesNo)

    End Sub
End Class
'Loop Until stcQuestionsInUse(0).strQuestion <> stcQuestionsInUse(1).strQuestion Or stcQuestionsInUse(0).strQuestion <> stcQuestionsInUse(2).strQuestion _
'    Or stcQuestionsInUse(0).strQuestion <> stcQuestionsInUse(3).strQuestion Or stcQuestionsInUse(0).strQuestion <> stcQuestionsInUse(4).strQuestion _
'    Or stcQuestionsInUse(0).strQuestion <> stcQuestionsInUse(5).strQuestion Or stcQuestionsInUse(0).strQuestion <> stcQuestionsInUse(6).strQuestion _
'    Or stcQuestionsInUse(0).strQuestion <> stcQuestionsInUse(7).strQuestion Or stcQuestionsInUse(0).strQuestion <> stcQuestionsInUse(8).strQuestion _
'    Or stcQuestionsInUse(0).strQuestion <> stcQuestionsInUse(9).strQuestion Or stcQuestionsInUse(0).strQuestion <> stcQuestionsInUse(10).strQuestion _
'    Or stcQuestionsInUse(0).strQuestion <> stcQuestionsInUse(11).strQuestion Or stcQuestionsInUse(0).strQuestion <> stcQuestionsInUse(12).strQuestion _
'    Or stcQuestionsInUse(0).strQuestion <> stcQuestionsInUse(13).strQuestion Or stcQuestionsInUse(0).strQuestion <> stcQuestionsInUse(14).strQuestion _
'    Or stcQuestionsInUse(1).strQuestion <> stcQuestionsInUse(0).strQuestion Or stcQuestionsInUse(1).strQuestion <> stcQuestionsInUse(2).strQuestion _
'    Or stcQuestionsInUse(1).strQuestion <> stcQuestionsInUse(3).strQuestion Or stcQuestionsInUse(1).strQuestion <> stcQuestionsInUse(4).strQuestion _
'    Or stcQuestionsInUse(1).strQuestion <> stcQuestionsInUse(5).strQuestion Or stcQuestionsInUse(1).strQuestion <> stcQuestionsInUse(6).strQuestion _
'    Or stcQuestionsInUse(1).strQuestion <> stcQuestionsInUse(7).strQuestion Or stcQuestionsInUse(1).strQuestion <> stcQuestionsInUse(8).strQuestion _
'    Or stcQuestionsInUse(1).strQuestion <> stcQuestionsInUse(9).strQuestion Or stcQuestionsInUse(1).strQuestion <> stcQuestionsInUse(10).strQuestion _
'    Or stcQuestionsInUse(1).strQuestion <> stcQuestionsInUse(11).strQuestion Or stcQuestionsInUse(1).strQuestion <> stcQuestionsInUse(12).strQuestion _
'    Or stcQuestionsInUse(1).strQuestion <> stcQuestionsInUse(13).strQuestion Or stcQuestionsInUse(1).strQuestion <> stcQuestionsInUse(14).strQuestion _
'    Or stcQuestionsInUse(2).strQuestion <> stcQuestionsInUse(1).strQuestion Or stcQuestionsInUse(2).strQuestion <> stcQuestionsInUse(0).strQuestion _
'    Or stcQuestionsInUse(2).strQuestion <> stcQuestionsInUse(3).strQuestion Or stcQuestionsInUse(2).strQuestion <> stcQuestionsInUse(4).strQuestion _
'    Or stcQuestionsInUse(2).strQuestion <> stcQuestionsInUse(5).strQuestion Or stcQuestionsInUse(2).strQuestion <> stcQuestionsInUse(6).strQuestion _
'    Or stcQuestionsInUse(2).strQuestion <> stcQuestionsInUse(7).strQuestion Or stcQuestionsInUse(2).strQuestion <> stcQuestionsInUse(8).strQuestion _
'    Or stcQuestionsInUse(2).strQuestion <> stcQuestionsInUse(9).strQuestion Or stcQuestionsInUse(2).strQuestion <> stcQuestionsInUse(10).strQuestion _
'    Or stcQuestionsInUse(2).strQuestion <> stcQuestionsInUse(11).strQuestion Or stcQuestionsInUse(2).strQuestion <> stcQuestionsInUse(12).strQuestion _
'    Or stcQuestionsInUse(2).strQuestion <> stcQuestionsInUse(13).strQuestion Or stcQuestionsInUse(2).strQuestion <> stcQuestionsInUse(14).strQuestion _
'    Or stcQuestionsInUse(3).strQuestion <> stcQuestionsInUse(1).strQuestion Or stcQuestionsInUse(3).strQuestion <> stcQuestionsInUse(2).strQuestion _
'    Or stcQuestionsInUse(3).strQuestion <> stcQuestionsInUse(0).strQuestion Or stcQuestionsInUse(3).strQuestion <> stcQuestionsInUse(4).strQuestion _
'    Or stcQuestionsInUse(3).strQuestion <> stcQuestionsInUse(5).strQuestion Or stcQuestionsInUse(3).strQuestion <> stcQuestionsInUse(6).strQuestion _
'    Or stcQuestionsInUse(3).strQuestion <> stcQuestionsInUse(7).strQuestion Or stcQuestionsInUse(3).strQuestion <> stcQuestionsInUse(8).strQuestion _
'    Or stcQuestionsInUse(3).strQuestion <> stcQuestionsInUse(9).strQuestion Or stcQuestionsInUse(3).strQuestion <> stcQuestionsInUse(10).strQuestion _
'    Or stcQuestionsInUse(3).strQuestion <> stcQuestionsInUse(11).strQuestion Or stcQuestionsInUse(3).strQuestion <> stcQuestionsInUse(12).strQuestion _
'    Or stcQuestionsInUse(3).strQuestion <> stcQuestionsInUse(13).strQuestion Or stcQuestionsInUse(3).strQuestion <> stcQuestionsInUse(14).strQuestion _
'    Or stcQuestionsInUse(4).strQuestion <> stcQuestionsInUse(1).strQuestion Or stcQuestionsInUse(4).strQuestion <> stcQuestionsInUse(2).strQuestion _
'    Or stcQuestionsInUse(4).strQuestion <> stcQuestionsInUse(3).strQuestion Or stcQuestionsInUse(4).strQuestion <> stcQuestionsInUse(0).strQuestion _
'    Or stcQuestionsInUse(4).strQuestion <> stcQuestionsInUse(5).strQuestion Or stcQuestionsInUse(4).strQuestion <> stcQuestionsInUse(6).strQuestion _
'    Or stcQuestionsInUse(4).strQuestion <> stcQuestionsInUse(7).strQuestion Or stcQuestionsInUse(4).strQuestion <> stcQuestionsInUse(8).strQuestion _
'    Or stcQuestionsInUse(4).strQuestion <> stcQuestionsInUse(9).strQuestion Or stcQuestionsInUse(4).strQuestion <> stcQuestionsInUse(10).strQuestion _
'    Or stcQuestionsInUse(4).strQuestion <> stcQuestionsInUse(11).strQuestion Or stcQuestionsInUse(4).strQuestion <> stcQuestionsInUse(12).strQuestion _
'    Or stcQuestionsInUse(4).strQuestion <> stcQuestionsInUse(13).strQuestion Or stcQuestionsInUse(4).strQuestion <> stcQuestionsInUse(14).strQuestion _
'    Or stcQuestionsInUse(5).strQuestion <> stcQuestionsInUse(1).strQuestion Or stcQuestionsInUse(5).strQuestion <> stcQuestionsInUse(2).strQuestion _
'    Or stcQuestionsInUse(5).strQuestion <> stcQuestionsInUse(3).strQuestion Or stcQuestionsInUse(5).strQuestion <> stcQuestionsInUse(4).strQuestion _
'    Or stcQuestionsInUse(5).strQuestion <> stcQuestionsInUse(0).strQuestion Or stcQuestionsInUse(5).strQuestion <> stcQuestionsInUse(6).strQuestion _
'    Or stcQuestionsInUse(5).strQuestion <> stcQuestionsInUse(7).strQuestion Or stcQuestionsInUse(5).strQuestion <> stcQuestionsInUse(8).strQuestion _
'    Or stcQuestionsInUse(5).strQuestion <> stcQuestionsInUse(9).strQuestion Or stcQuestionsInUse(5).strQuestion <> stcQuestionsInUse(10).strQuestion _
'    Or stcQuestionsInUse(5).strQuestion <> stcQuestionsInUse(11).strQuestion Or stcQuestionsInUse(5).strQuestion <> stcQuestionsInUse(12).strQuestion _
'    Or stcQuestionsInUse(5).strQuestion <> stcQuestionsInUse(13).strQuestion Or stcQuestionsInUse(5).strQuestion <> stcQuestionsInUse(14).strQuestion _
'    Or stcQuestionsInUse(6).strQuestion <> stcQuestionsInUse(1).strQuestion Or stcQuestionsInUse(6).strQuestion <> stcQuestionsInUse(2).strQuestion _
'    Or stcQuestionsInUse(6).strQuestion <> stcQuestionsInUse(3).strQuestion Or stcQuestionsInUse(6).strQuestion <> stcQuestionsInUse(4).strQuestion _
'    Or stcQuestionsInUse(6).strQuestion <> stcQuestionsInUse(5).strQuestion Or stcQuestionsInUse(6).strQuestion <> stcQuestionsInUse(0).strQuestion _
'    Or stcQuestionsInUse(6).strQuestion <> stcQuestionsInUse(7).strQuestion Or stcQuestionsInUse(6).strQuestion <> stcQuestionsInUse(8).strQuestion _
'    Or stcQuestionsInUse(6).strQuestion <> stcQuestionsInUse(9).strQuestion Or stcQuestionsInUse(6).strQuestion <> stcQuestionsInUse(10).strQuestion _
'    Or stcQuestionsInUse(6).strQuestion <> stcQuestionsInUse(11).strQuestion Or stcQuestionsInUse(6).strQuestion <> stcQuestionsInUse(12).strQuestion _
'    Or stcQuestionsInUse(6).strQuestion <> stcQuestionsInUse(13).strQuestion Or stcQuestionsInUse(6).strQuestion <> stcQuestionsInUse(14).strQuestion _
'    Or stcQuestionsInUse(7).strQuestion <> stcQuestionsInUse(1).strQuestion Or stcQuestionsInUse(7).strQuestion <> stcQuestionsInUse(2).strQuestion _
'    Or stcQuestionsInUse(7).strQuestion <> stcQuestionsInUse(3).strQuestion Or stcQuestionsInUse(7).strQuestion <> stcQuestionsInUse(4).strQuestion _
'    Or stcQuestionsInUse(7).strQuestion <> stcQuestionsInUse(5).strQuestion Or stcQuestionsInUse(7).strQuestion <> stcQuestionsInUse(6).strQuestion _
'    Or stcQuestionsInUse(7).strQuestion <> stcQuestionsInUse(0).strQuestion Or stcQuestionsInUse(7).strQuestion <> stcQuestionsInUse(8).strQuestion _
'    Or stcQuestionsInUse(7).strQuestion <> stcQuestionsInUse(9).strQuestion Or stcQuestionsInUse(7).strQuestion <> stcQuestionsInUse(10).strQuestion _
'    Or stcQuestionsInUse(7).strQuestion <> stcQuestionsInUse(11).strQuestion Or stcQuestionsInUse(7).strQuestion <> stcQuestionsInUse(12).strQuestion _
'    Or stcQuestionsInUse(7).strQuestion <> stcQuestionsInUse(13).strQuestion Or stcQuestionsInUse(7).strQuestion <> stcQuestionsInUse(14).strQuestion _
'    Or stcQuestionsInUse(8).strQuestion <> stcQuestionsInUse(1).strQuestion Or stcQuestionsInUse(8).strQuestion <> stcQuestionsInUse(2).strQuestion _
'    Or stcQuestionsInUse(8).strQuestion <> stcQuestionsInUse(3).strQuestion Or stcQuestionsInUse(8).strQuestion <> stcQuestionsInUse(4).strQuestion _
'    Or stcQuestionsInUse(8).strQuestion <> stcQuestionsInUse(5).strQuestion Or stcQuestionsInUse(8).strQuestion <> stcQuestionsInUse(6).strQuestion _
'    Or stcQuestionsInUse(8).strQuestion <> stcQuestionsInUse(7).strQuestion Or stcQuestionsInUse(8).strQuestion <> stcQuestionsInUse(0).strQuestion _
'    Or stcQuestionsInUse(8).strQuestion <> stcQuestionsInUse(9).strQuestion Or stcQuestionsInUse(8).strQuestion <> stcQuestionsInUse(10).strQuestion _
'    Or stcQuestionsInUse(8).strQuestion <> stcQuestionsInUse(11).strQuestion Or stcQuestionsInUse(8).strQuestion <> stcQuestionsInUse(12).strQuestion _
'    Or stcQuestionsInUse(8).strQuestion <> stcQuestionsInUse(13).strQuestion Or stcQuestionsInUse(8).strQuestion <> stcQuestionsInUse(14).strQuestion _
'    Or stcQuestionsInUse(9).strQuestion <> stcQuestionsInUse(1).strQuestion Or stcQuestionsInUse(9).strQuestion <> stcQuestionsInUse(2).strQuestion _
'    Or stcQuestionsInUse(9).strQuestion <> stcQuestionsInUse(3).strQuestion Or stcQuestionsInUse(9).strQuestion <> stcQuestionsInUse(4).strQuestion _
'    Or stcQuestionsInUse(9).strQuestion <> stcQuestionsInUse(5).strQuestion Or stcQuestionsInUse(9).strQuestion <> stcQuestionsInUse(6).strQuestion _
'    Or stcQuestionsInUse(9).strQuestion <> stcQuestionsInUse(7).strQuestion Or stcQuestionsInUse(9).strQuestion <> stcQuestionsInUse(8).strQuestion _
'    Or stcQuestionsInUse(9).strQuestion <> stcQuestionsInUse(0).strQuestion Or stcQuestionsInUse(9).strQuestion <> stcQuestionsInUse(10).strQuestion _
'    Or stcQuestionsInUse(9).strQuestion <> stcQuestionsInUse(11).strQuestion Or stcQuestionsInUse(9).strQuestion <> stcQuestionsInUse(12).strQuestion _
'    Or stcQuestionsInUse(9).strQuestion <> stcQuestionsInUse(13).strQuestion Or stcQuestionsInUse(9).strQuestion <> stcQuestionsInUse(14).strQuestion _
'    Or stcQuestionsInUse(10).strQuestion <> stcQuestionsInUse(1).strQuestion Or stcQuestionsInUse(10).strQuestion <> stcQuestionsInUse(2).strQuestion _
'    Or stcQuestionsInUse(10).strQuestion <> stcQuestionsInUse(3).strQuestion Or stcQuestionsInUse(10).strQuestion <> stcQuestionsInUse(4).strQuestion _
'    Or stcQuestionsInUse(10).strQuestion <> stcQuestionsInUse(5).strQuestion Or stcQuestionsInUse(10).strQuestion <> stcQuestionsInUse(6).strQuestion _
'    Or stcQuestionsInUse(10).strQuestion <> stcQuestionsInUse(7).strQuestion Or stcQuestionsInUse(10).strQuestion <> stcQuestionsInUse(8).strQuestion _
'    Or stcQuestionsInUse(10).strQuestion <> stcQuestionsInUse(9).strQuestion Or stcQuestionsInUse(10).strQuestion <> stcQuestionsInUse(0).strQuestion _
'    Or stcQuestionsInUse(10).strQuestion <> stcQuestionsInUse(11).strQuestion Or stcQuestionsInUse(10).strQuestion <> stcQuestionsInUse(12).strQuestion _
'    Or stcQuestionsInUse(10).strQuestion <> stcQuestionsInUse(13).strQuestion Or stcQuestionsInUse(10).strQuestion <> stcQuestionsInUse(14).strQuestion _
'    Or stcQuestionsInUse(11).strQuestion <> stcQuestionsInUse(1).strQuestion Or stcQuestionsInUse(11).strQuestion <> stcQuestionsInUse(2).strQuestion _
'    Or stcQuestionsInUse(11).strQuestion <> stcQuestionsInUse(3).strQuestion Or stcQuestionsInUse(11).strQuestion <> stcQuestionsInUse(4).strQuestion _
'    Or stcQuestionsInUse(11).strQuestion <> stcQuestionsInUse(5).strQuestion Or stcQuestionsInUse(11).strQuestion <> stcQuestionsInUse(6).strQuestion _
'    Or stcQuestionsInUse(11).strQuestion <> stcQuestionsInUse(7).strQuestion Or stcQuestionsInUse(11).strQuestion <> stcQuestionsInUse(8).strQuestion _
'    Or stcQuestionsInUse(11).strQuestion <> stcQuestionsInUse(9).strQuestion Or stcQuestionsInUse(11).strQuestion <> stcQuestionsInUse(10).strQuestion _
'    Or stcQuestionsInUse(11).strQuestion <> stcQuestionsInUse(0).strQuestion Or stcQuestionsInUse(11).strQuestion <> stcQuestionsInUse(12).strQuestion _
'    Or stcQuestionsInUse(11).strQuestion <> stcQuestionsInUse(13).strQuestion Or stcQuestionsInUse(11).strQuestion <> stcQuestionsInUse(14).strQuestion _
'    Or stcQuestionsInUse(12).strQuestion <> stcQuestionsInUse(1).strQuestion Or stcQuestionsInUse(12).strQuestion <> stcQuestionsInUse(2).strQuestion _
'    Or stcQuestionsInUse(12).strQuestion <> stcQuestionsInUse(3).strQuestion Or stcQuestionsInUse(12).strQuestion <> stcQuestionsInUse(4).strQuestion _
'    Or stcQuestionsInUse(12).strQuestion <> stcQuestionsInUse(5).strQuestion Or stcQuestionsInUse(12).strQuestion <> stcQuestionsInUse(6).strQuestion _
'    Or stcQuestionsInUse(12).strQuestion <> stcQuestionsInUse(7).strQuestion Or stcQuestionsInUse(12).strQuestion <> stcQuestionsInUse(8).strQuestion _
'    Or stcQuestionsInUse(12).strQuestion <> stcQuestionsInUse(9).strQuestion Or stcQuestionsInUse(12).strQuestion <> stcQuestionsInUse(10).strQuestion _
'    Or stcQuestionsInUse(12).strQuestion <> stcQuestionsInUse(11).strQuestion Or stcQuestionsInUse(12).strQuestion <> stcQuestionsInUse(0).strQuestion _
'    Or stcQuestionsInUse(12).strQuestion <> stcQuestionsInUse(13).strQuestion Or stcQuestionsInUse(12).strQuestion <> stcQuestionsInUse(14).strQuestion _
'    Or stcQuestionsInUse(13).strQuestion <> stcQuestionsInUse(1).strQuestion Or stcQuestionsInUse(13).strQuestion <> stcQuestionsInUse(2).strQuestion _
'    Or stcQuestionsInUse(13).strQuestion <> stcQuestionsInUse(3).strQuestion Or stcQuestionsInUse(13).strQuestion <> stcQuestionsInUse(4).strQuestion _
'    Or stcQuestionsInUse(13).strQuestion <> stcQuestionsInUse(5).strQuestion Or stcQuestionsInUse(13).strQuestion <> stcQuestionsInUse(6).strQuestion _
'    Or stcQuestionsInUse(13).strQuestion <> stcQuestionsInUse(7).strQuestion Or stcQuestionsInUse(13).strQuestion <> stcQuestionsInUse(8).strQuestion _
'    Or stcQuestionsInUse(13).strQuestion <> stcQuestionsInUse(9).strQuestion Or stcQuestionsInUse(13).strQuestion <> stcQuestionsInUse(10).strQuestion _
'    Or stcQuestionsInUse(13).strQuestion <> stcQuestionsInUse(11).strQuestion Or stcQuestionsInUse(13).strQuestion <> stcQuestionsInUse(12).strQuestion _
'    Or stcQuestionsInUse(13).strQuestion <> stcQuestionsInUse(13).strQuestion Or stcQuestionsInUse(13).strQuestion <> stcQuestionsInUse(14).strQuestion _
'    Or stcQuestionsInUse(14).strQuestion <> stcQuestionsInUse(1).strQuestion Or stcQuestionsInUse(14).strQuestion <> stcQuestionsInUse(2).strQuestion _
'    Or stcQuestionsInUse(14).strQuestion <> stcQuestionsInUse(3).strQuestion Or stcQuestionsInUse(14).strQuestion <> stcQuestionsInUse(4).strQuestion _
'    Or stcQuestionsInUse(14).strQuestion <> stcQuestionsInUse(5).strQuestion Or stcQuestionsInUse(14).strQuestion <> stcQuestionsInUse(6).strQuestion _
'    Or stcQuestionsInUse(14).strQuestion <> stcQuestionsInUse(7).strQuestion Or stcQuestionsInUse(14).strQuestion <> stcQuestionsInUse(8).strQuestion _
'    Or stcQuestionsInUse(14).strQuestion <> stcQuestionsInUse(9).strQuestion Or stcQuestionsInUse(14).strQuestion <> stcQuestionsInUse(10).strQuestion _
'    Or stcQuestionsInUse(14).strQuestion <> stcQuestionsInUse(11).strQuestion Or stcQuestionsInUse(14).strQuestion <> stcQuestionsInUse(12).strQuestion _
'    Or stcQuestionsInUse(14).strQuestion <> stcQuestionsInUse(13).strQuestion Or stcQuestionsInUse(14).strQuestion <> stcQuestionsInUse(0).strQuestion _
'    Or stcQuestionsInUse(0).strQuestion <> Nothing Or stcQuestionsInUse(1).strQuestion <> Nothing Or stcQuestionsInUse(2).strQuestion <> Nothing _
'    Or stcQuestionsInUse(3).strQuestion <> Nothing Or stcQuestionsInUse(4).strQuestion <> Nothing Or stcQuestionsInUse(5).strQuestion <> Nothing _
'    Or stcQuestionsInUse(6).strQuestion <> Nothing Or stcQuestionsInUse(7).strQuestion <> Nothing Or stcQuestionsInUse(8).strQuestion <> Nothing _
'    Or stcQuestionsInUse(9).strQuestion <> Nothing Or stcQuestionsInUse(10).strQuestion <> Nothing _
'    Or stcQuestionsInUse(11).strQuestion <> Nothing Or stcQuestionsInUse(12).strQuestion <> Nothing Or stcQuestionsInUse(13).strQuestion <> Nothing _
'    Or stcQuestionsInUse(14).strQuestion <> Nothing

