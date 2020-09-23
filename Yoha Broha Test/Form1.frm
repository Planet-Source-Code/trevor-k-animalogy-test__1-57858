VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4920
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6975
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4920
   ScaleWidth      =   6975
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   495
      Left            =   5760
      TabIndex        =   10
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Start Over"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   9
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF0000&
      Caption         =   "Next"
      Height          =   495
      Left            =   0
      TabIndex        =   8
      Top             =   4440
      Width           =   1215
   End
   Begin VB.OptionButton Option6 
      Caption         =   "Option6"
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   3720
      Width           =   6735
   End
   Begin VB.OptionButton Option5 
      Caption         =   "Option5"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   3240
      Width           =   6735
   End
   Begin VB.OptionButton Option4 
      Caption         =   "Option4"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   2760
      Width           =   6735
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Option3"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   2280
      Width           =   6735
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Option2"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   6735
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   6735
   End
   Begin VB.Frame Frame1 
      Caption         =   "Choose One"
      Height          =   3255
      Left            =   0
      TabIndex        =   1
      Top             =   1080
      Width           =   6975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private xx As Integer

Private Sub Command1_Click()
If xx = 1 Then
    If Option1.Value = True Then Call two
    If Option2.Value = True Then Call three
    GoTo ED
    End If
If xx = 2 Then
    If Option1.Value = True Then Call seven
    If Option2.Value = True Then Call three
    If Option3.Value = True Then Call five
    GoTo ED
    End If
If xx = 3 Then
    If Option1.Value = True Then Call five
    If Option2.Value = True Then Call four
    GoTo ED
    End If
If xx = 4 Then
    If Option1.Value = True Then Call ten
    If Option2.Value = True Then Call eight
    If Option3.Value = True Then Call nine
    GoTo ED
    End If
If xx = 5 Then
    If Option1.Value = True Then Call seven
    If Option2.Value = True Then Call six
    If Option3.Value = True Then Call four
    GoTo ED
    End If
If xx = 6 Then
    If Option1.Value = True Then Call eight
    If Option2.Value = True Then Call nine
    If Option3.Value = True Then Call seven
    If Option4.Value = True Then Call seven
    GoTo ED
    End If
If xx = 7 Then
    If Option1.Value = True Then Call eight
    If Option2.Value = True Then Call ten
    GoTo ED
    End If
If xx = 8 Then
    If Option1.Value = True Then Call ten
    If Option2.Value = True Then Call nine
    If Option3.Value = True Then Call p
    GoTo ED
    End If
If xx = 9 Then
    If Option1.Value = True Then Call y
    If Option2.Value = True Then Call a
    If Option3.Value = True Then Call eleven
    If Option4.Value = True Then Call u
    GoTo ED
    End If
If xx = 10 Then
    If Option1.Value = True Then Call eleven
    If Option2.Value = True Then Call fifteen
    If Option3.Value = True Then Call fourteen
    If Option4.Value = True Then Call twelve
    GoTo ED
    End If
If xx = 11 Then
    If Option1.Value = True Then Call b
    If Option2.Value = True Then Call twelve
    If Option3.Value = True Then Call r
    If Option4.Value = True Then Call twentyfive
    GoTo ED
    End If
If xx = 12 Then
    If Option1.Value = True Then Call thirteen
    If Option2.Value = True Then Call eighteen
    If Option3.Value = True Then Call fourteen
    GoTo ED
    End If
If xx = 13 Then
    If Option1.Value = True Then Call fifteen
    If Option2.Value = True Then Call g
    If Option3.Value = True Then Call d
    GoTo ED
    End If
If xx = 14 Then
    If Option1.Value = True Then Call seventeen
    If Option2.Value = True Then Call sixteen
    If Option3.Value = True Then Call twentytwo
    GoTo ED
    End If
If xx = 15 Then
    If Option1.Value = True Then Call c
    If Option2.Value = True Then Call f
    GoTo ED
    End If
If xx = 16 Then
    If Option1.Value = True Then Call seventeen
    If Option2.Value = True Then Call twentytwo
    GoTo ED
    End If
If xx = 17 Then
    If Option1.Value = True Then Call e
    If Option2.Value = True Then Call twenty
    If Option3.Value = True Then Call l
    If Option4.Value = True Then Call twentyfour
    GoTo ED
    End If
If xx = 18 Then
    If Option1.Value = True Then Call w
    If Option2.Value = True Then Call i
    If Option3.Value = True Then Call nineteen
    GoTo ED
    End If
If xx = 19 Then
    If Option1.Value = True Then Call h
    If Option2.Value = True Then Call twentyone
    If Option3.Value = True Then Call twenty
    GoTo ED
    End If
If xx = 20 Then
    If Option1.Value = True Then Call j
    If Option2.Value = True Then Call twentythree
    If Option3.Value = True Then Call m
    If Option4.Value = True Then Call e
    GoTo ED
    End If
If xx = 21 Then
    If Option1.Value = True Then Call twentythree
    If Option2.Value = True Then Call u
    If Option3.Value = True Then Call l
    GoTo ED
    End If
If xx = 22 Then
    If Option1.Value = True Then Call x
    If Option2.Value = True Then Call n
    If Option3.Value = True Then Call nineteen
    GoTo ED
    End If
If xx = 23 Then
    If Option1.Value = True Then Call k
    If Option2.Value = True Then Call twentyfour
    If Option3.Value = True Then Call q
    If Option4.Value = True Then Call twentyfour
    GoTo ED
    End If
If xx = 24 Then
    If Option1.Value = True Then Call t
    If Option2.Value = True Then Call s
    If Option3.Value = True Then Call v
    If Option4.Value = True Then Call twentyfive
    If Option5.Value = True Then Call twentysix
    GoTo ED
    End If
If xx = 25 Then
    If Option1.Value = True Then Call z
    If Option2.Value = True Then Call ab
    If Option3.Value = True Then Call ac
    If Option4.Value = True Then Call ad
    If Option5.Value = True Then Call ae
    If Option6.Value = True Then Call y
    GoTo ED
    End If
If xx = 26 Then
    If Option1.Value = True Then Call twentyseven
    If Option2.Value = True Then Call twentyfive
    If Option3.Value = True Then Call ae
    If Option4.Value = True Then Call h
    GoTo ED
    End If
If xx = 27 Then
    If Option1.Value = True Then Call twentyfive
    If Option2.Value = True Then Call f
    If Option3.Value = True Then Call af
    If Option4.Value = True Then Call ag
    GoTo ED
    End If
ED:
End Sub

Private Sub Command2_Click()
Call Form_Load

End Sub
Private Sub Command3_Click()
End
End Sub
Private Sub Reset()
Option1.Caption = "a"
Option2.Caption = "b"
Option3.Caption = "c"
Option4.Caption = "d"
Option5.Caption = "e"
Option6.Caption = "f"
Option1.Enabled = False
Option2.Enabled = False
Option3.Enabled = False
Option4.Enabled = False
Option5.Enabled = False
Option6.Enabled = False
Label1.Caption = ""
End Sub
Private Sub Form_Load()
Call Reset
Call one
End Sub
Private Sub one()
xx = 1
Label1.Caption = "Would you Classify yourself as someone who is involved in sports? (Ex. You take soccer)"
Option1.Enabled = True
Option2.Enabled = True
Option1.Caption = "Yes"
Option2.Caption = "No"
End Sub
Private Sub two()
Call Reset
xx = 2
Label1.Caption = "You would classify yourself as:"
Option1.Enabled = True
Option2.Enabled = True
Option3.Enabled = True
Option1.Caption = "A nerd/loser"
Option2.Caption = "A bully"
Option3.Caption = "An average guy/girl"
End Sub
Private Sub three()
Call Reset
xx = 3
Label1.Caption = "Would you accept money from a stranger?"
Option1.Enabled = True
Option2.Enabled = True
Option1.Caption = "Yes"
Option2.Caption = "No"
End Sub
Private Sub four()
Call Reset
xx = 4
Label1.Caption = "If you were a tree, what kind of tree would you be?"
Option1.Enabled = True
Option2.Enabled = True
Option3.Enabled = True
Option1.Caption = "Oak"
Option2.Caption = "Weeping Willow"
Option3.Caption = "Fir Tree"
End Sub
Private Sub five()
Call Reset
xx = 5
Label1.Caption = "A movie you want to see closed down. What will you do?"
Option1.Enabled = True
Option2.Enabled = True
Option3.Enabled = True
Option1.Caption = "Go home"
Option2.Caption = "Go see a different movie"
Option3.Caption = "Go do something else"
End Sub
Private Sub six()
Call Reset
xx = 6
Label1.Caption = "A mysterious hooded stranger arrives at your house. You are most likely to:"
Option1.Enabled = True
Option2.Enabled = True
Option3.Enabled = True
Option4.Enabled = True
Option1.Caption = "Greet him/her"
Option2.Caption = "Hide and make them think you're not home"
Option3.Caption = "Call the cops"
Option4.Caption = "Yell through the door I'VE GOT A GUN"
End Sub
Private Sub seven()
Call Reset
xx = 7
Label1.Caption = "Would you call yourself strong (phisically or mentally Which ever you want?"
Option1.Enabled = True
Option2.Enabled = True
Option1.Caption = "Yes"
Option2.Caption = "No"
End Sub
Private Sub eight()
Call Reset
xx = 8
Label1.Caption = "There is an earthquake. What will you do?"
Option1.Enabled = True
Option2.Enabled = True
Option3.Enabled = True
Option1.Caption = "Go out and help anyone in need"
Option2.Caption = "Go to a safe place with family and friends"
Option3.Caption = "Panic and wait for someone to help you"
End Sub
Private Sub nine()
Call Reset
xx = 9
Label1.Caption = "You buy ice, and a kid comes up and tells you he's really hungry and hasn't eaten in days. Your reaction will be:"
Option1.Enabled = True
Option2.Enabled = True
Option3.Enabled = True
Option4.Enabled = True
Option1.Caption = "To stick the ice cream in his face and run away"
Option2.Caption = "To ignore him and eat your ice cream"
Option3.Caption = "To give him you ice cream"
Option4.Caption = "You think he is lying"
End Sub
Private Sub ten()
Call Reset
xx = 10
Label1.Caption = "Your friend chalanges you to a game.  What will your answer be?"
Option1.Enabled = True
Option2.Enabled = True
Option3.Enabled = True
Option4.Enabled = True
Option1.Caption = "Bring it on!"
Option2.Caption = "Fine."
Option3.Caption = "Leave me alone!"
Option4.Caption = "Sorry, but I really don't feel like it."
End Sub
Private Sub eleven()
Call Reset
xx = 11
Label1.Caption = "You are going to take a quiz (like the one in school). Your score is most likely to be:"
Option1.Enabled = True
Option2.Enabled = True
Option3.Enabled = True
Option4.Enabled = True
Option1.Caption = "Perfect"
Option2.Caption = "Enough to pass"
Option3.Caption = "A failing mark. I don;t actually care anyways."
Option4.Caption = "I don't know"
End Sub
Private Sub twelve()
Call Reset
xx = 12
Label1.Caption = "How many friends do you have?"
Option1.Enabled = True
Option2.Enabled = True
Option3.Enabled = True
Option1.Caption = "One Only"
Option2.Caption = "About three to five"
Option3.Caption = "A LOT!"
End Sub
Private Sub thirteen()
Call Reset
xx = 13
Label1.Caption = "Are you constantly..."
Option1.Enabled = True
Option2.Enabled = True
Option3.Enabled = True
Option1.Caption = "Bullied?"
Option2.Caption = "Feared by others?"
Option3.Caption = "Idolized?"
End Sub
Private Sub fourteen()
Call Reset
xx = 14
Label1.Caption = "The food you eat is usually:"
Option1.Enabled = True
Option2.Enabled = True
Option3.Enabled = True
Option1.Caption = "Cheap"
Option2.Caption = "Expensive"
Option3.Caption = "Both! I love food!"
End Sub
Private Sub fifteen()
Call Reset
xx = 15
Label1.Caption = "Do you belive in the saying, An apple a day keeps the doctor away?"
Option1.Enabled = True
Option2.Enabled = True
Option1.Caption = "Yes"
Option2.Caption = "No"
End Sub
Private Sub sixteen()
Call Reset
xx = 16
Label1.Caption = "Do you belive in true love?"
Option1.Enabled = True
Option2.Enabled = True
Option1.Caption = "Yes"
Option2.Caption = "No"
End Sub
Private Sub seventeen()
Call Reset
xx = 17
Label1.Caption = "You see your crush.  You are most likely to:"
Option1.Enabled = True
Option2.Enabled = True
Option3.Enabled = True
Option4.Enabled = True
Option1.Caption = "Flirt with him/her"
Option2.Caption = "Run away"
Option3.Caption = "Act mean to him/her"
Option4.Caption = "Act normal"
End Sub
Private Sub eighteen()
Call Reset
xx = 18
Label1.Caption = "One of your bad traits (among these) is are:"
Option1.Enabled = True
Option2.Enabled = True
Option3.Enabled = True
Option1.Caption = "Being noisy/talkative"
Option2.Caption = "None at all. I'm perfect!"
Option3.Caption = "Being shy"
End Sub
Private Sub nineteen()
Call Reset
xx = 19
Label1.Caption = "What do you think of nerds/geeks/smart people?"
Option1.Enabled = True
Option2.Enabled = True
Option3.Enabled = True
Option1.Caption = "They're okay"
Option2.Caption = "EWWWWW!!!!!!"
Option3.Caption = "There Know-it-alls"
End Sub
Private Sub twenty()
Call Reset
xx = 20
Label1.Caption = "Someone asks you: What do you know about the Industrial Revolution?. Your initial response would be:"
Option1.Enabled = True
Option2.Enabled = True
Option3.Enabled = True
Option4.Enabled = True
Option1.Caption = "Yawn..."
Option2.Caption = "Oh, I knowx all about it!"
Option3.Caption = "Umm... I think I remember it from somewhere..."
Option4.Caption = "Why do you want to know!?"
End Sub
Private Sub twentyone()
Call Reset
xx = 21
Label1.Caption = "Someone goes up to you and says: Did you know that eating glue makes your brain stronger? Your initial response would be:"
Option1.Enabled = True
Option2.Enabled = True
Option3.Enabled = True
Option1.Caption = "Really?!"
Option2.Caption = "Yeah, right!"
Option3.Caption = "That's nice"
End Sub
Private Sub twentytwo()
Call Reset
xx = 22
Label1.Caption = "If you were in a human food chain, where would you put yourself?"
Option1.Enabled = True
Option2.Enabled = True
Option3.Enabled = True
Option1.Caption = "At the bottom... I'm low-class...."
Option2.Caption = "At the very top! I'm the predator of all preditors!"
Option3.Caption = "Somewhere in the middle, I guess..."
End Sub
Private Sub twentythree()
Call Reset
xx = 23
Label1.Caption = "Would you call yourself hardworking?"
Option1.Enabled = True
Option2.Enabled = True
Option3.Enabled = True
Option4.Enabled = True
Option1.Caption = "Yes"
Option2.Caption = "No"
Option3.Caption = "Uhhh... Sometimes"
Option4.Caption = "no, more then yes"
End Sub
Private Sub twentyfour()
Call Reset
xx = 24
Label1.Caption = "If you had a special power, which power would you want to have?"
Option1.Enabled = True
Option2.Enabled = True
Option3.Enabled = True
Option4.Enabled = True
Option5.Enabled = True
Option1.Caption = "The power to be invisable"
Option2.Caption = "The power to be the greatest person on Earth!"
Option3.Caption = "The power to be the fastest thing alive!"
Option4.Caption = "I don't need special powers."
Option5.Caption = "The power to teleport"
End Sub
Private Sub twentyfive()
Call Reset
xx = 25
Label1.Caption = "You are known as:"
Option1.Enabled = True
Option2.Enabled = True
Option3.Enabled = True
Option4.Enabled = True
Option5.Enabled = True
Option6.Enabled = True
Option1.Caption = "A loner"
Option2.Caption = "A good friend to all"
Option3.Caption = "A food expert"
Option4.Caption = "A marjory fun person"
Option5.Caption = "A person with a split personality"
Option6.Caption = "A wierd person"
End Sub
Private Sub twentysix()
Call Reset
xx = 26
Label1.Caption = "Which of these is a great attribute about you?"
Option1.Enabled = True
Option2.Enabled = True
Option3.Enabled = True
Option4.Enabled = True
Option1.Caption = "You are Flexable"
Option2.Caption = "You are strong (Mentally or phisically)"
Option3.Caption = "You are retarded"
Option4.Caption = "You are normal"
End Sub
Private Sub twentyseven()
Call Reset
xx = 27
Label1.Caption = "Who do you hangout with?"
Option1.Enabled = True
Option2.Enabled = True
Option3.Enabled = True
Option4.Enabled = True
Option1.Caption = "Smart people"
Option2.Caption = "Gangsters"
Option3.Caption = "Weird people"
Option4.Caption = "Every one"
End Sub
Private Sub a()
pa$ = "Peach pig" & vbCrLf & vbCrLf & "You are a big, lazy slob. Yet, you have a lot of strength but are afraid to use it." & vbCrLf & "Soul Mate: Magenta Squirrel" & vbCrLf & "In conflict with: Ocre and gray Dolphin"
MsgBox pa$
End Sub
Private Sub b()
pb$ = "Blue Fox" & vbCrLf & vbCrLf & "You are a born leader who cannot resest a challange. You are driven to excel and a perfectionist." & vbCrLf & "Soul Mate: Yellow Trout" & vbCrLf & "In comflict with: Indigo Beaver"
MsgBox pb$
End Sub
Private Sub c()
pc$ = "Pink Sloth" & vbCrLf & vbCrLf & "You are an outcast; A follower, socially inept. You are the lowest of all animology, mainly because you smell of overripe fruit." & vbCrLf & "Soul Mate: Silver Badger" & vbCrLf & "In conflict with: Tawny Mouse"
MsgBox pc$
End Sub
Private Sub d()
pd$ = "Teal Cat" & vbCrLf & vbCrLf & "You are as swift as a ninja. You can be soft-hearted and cruel at the same time, and extremely hard to please." & vbCrLf & "Soul Mate: Beige Raccoon" & vbCrLf & "In conflict with: Red Jaguar"
MsgBox pd$
End Sub
Private Sub e()
pe$ = "Yellow Trout" & vbCrLf & vbCrLf & "You are extremely self-centered and only think of yourself. You are also very whiny and annoying to others, but you are able to get away with it. Aside from that, you have slow reactions, except when you are in love." & vbCrLf & "Soul Mate: Blue fox" & vbCrLf & "In conflict with: Green Puppy"
MsgBox pe$
End Sub
Private Sub f()
pf$ = "Blue Baboon" & vbCrLf & vbCrLf & "An aggressive animology, especially to your enemies. You also have quick reflexes, but it takes you an unnaturally long time to remember something." & vbCrLf & "Soul Mate: Purple Bat" & vbCrLf & "In conflict with: Red Weasel"
MsgBox pf$
End Sub
Private Sub g()
pg$ = "Red Weasel" & vbCrLf & vbCrLf & "You are an extreme know-it-all, even if you're not perfect. You are also noisy, but far from hardworking." & vbCrLf & "Soul Mate: Brown Iguana" & vbCrLf & "In conflict with: Blue baboon"
MsgBox pg$
End Sub
Private Sub h()
ph$ = "Silver Badger" & vbCrLf & vbCrLf & "You are very loyal to others and hardworking. You are a good friend, even if others find you sometimes a bit boring." & vbCrLf & "Soul Mate: Pink Sloth" & vbCrLf & "In conflict with: White Tiger"
MsgBox ph$
End Sub
Private Sub i()
pi$ = "Orange Snake" & vbCrLf & vbCrLf & "Like the snake itself, you are cunning and verry boastful. You love to be the best, even if you're not, but you try hard enough." & vbCrLf & "Soul Mate: Periwinkle Cow" & vbCrLf & "In conflict with: Gray Chicken"
MsgBox pi$
End Sub
Private Sub j()
pj$ = "Green Puppy" & vbCrLf & vbCrLf & "You are very patient, but often moody. You like to strike when least expected, even if you're not exactly what one would call observant." & vbCrLf & "Soul Mate: Fuchsia Possum" & vbCrLf & "In conflict with: Yellow Trout"
MsgBox pj$
End Sub
Private Sub k()
pk$ = "Indigo Beaver" & vbCrLf & vbCrLf & "You are very resourceful and patient, but you are often naïve and fall for others' traps. Luckily, you get out of most of them." & vbCrLf & "Soul Mate: Lavender Cheetah" & vbCrLf & "In conflict with: Blue Fox"
MsgBox pk$
End Sub
Private Sub l()
pl$ = "Purple Bat" & vbCrLf & vbCrLf & "You are often blind to the world, and because of this, you use your heart to sense things. Yet, you are feared by other animologies, because you don't always listen to what your heart has to say, and you end up doing cruel things." & vbCrLf & "Soul Mate: Blue Baboon" & vbCrLf & "In conflict with: Golden Lion"
MsgBox pl$
End Sub
Private Sub m()
pm$ = "Black Sheep" & vbCrLf & vbCrLf & "You are different from other animologies. Even if some people think you are weak and are a disgrace, your heart always remains strong and pure." & vbCrLf & "Soul Mate: Maroon Panda" & vbCrLf & "In conflict with: Magenta Squirrel"
MsgBox pm$
End Sub
Private Sub n()
pn$ = "Golden Lion" & vbCrLf & vbCrLf & "You consider yourself as the king of all animologies and have a lot of dignity. But despite this statement, you are lazy and often end up embarrassing yourself." & vbCrLf & "Soul Mate: White Tiger" & vbCrLf & "In conflict with: Silver Badger"
MsgBox pn$
End Sub
Private Sub o()
po$ = "Brown Iguana" & vbCrLf & vbCrLf & "You are quiet and often stay in one place pondering your next move. You are not a threat to other animologies, even if your strikes are quite lethal." & vbCrLf & "Soul Mate: Red Weasel" & vbCrLf & "In Conflict with: Fuchsia Possum"
MsgBox po$
End Sub
Private Sub p()
pp$ = "Gray Chicken" & vbCrLf & vbCrLf & "You are cowardly, and quick to react. When danger strikes, you immediately run away, only thinking of yourself." & vbCrLf & "Soul Mate: Tawny Mouse" & vbCrLf & "In conflict with: Orange Snake"
MsgBox pp$
End Sub
Private Sub q()
pq$ = "White Tiger" & vbCrLf & vbCrLf & "You are unique; A graceful yet cunning animology. Yet, even with your strength, you are very easy prey to others." & vbCrLf & "Soul Mate: Golden Lion" & vbCrLf & "In conflict with: Silver Badger"
MsgBox pq$
End Sub
Private Sub r()
pr$ = "Scarlet Bear" & vbCrLf & vbCrLf & "You are constantly a bully to others, but wit is your ultimate weakness. Another disadvantage is that you're slow, but make up with your great amount of strength." & vbCrLf & "Soul Mate: Saffron Rabbit" & vbCrLf & "In conflict with: Periwinkle Cow"
MsgBox pr$
End Sub
Private Sub s()
ps$ = "Red Jaguar" & vbCrLf & vbCrLf & "Extremely quick, but you are easy to trick. All of your senses are quite strong, but you get caught, mainly because of too much pride in yourself which causes insecurity." & vbCrLf & "Soul Mate: Tan Giraffe" & vbCrLf & "In conflict with: Teal Cat"
MsgBox ps$
End Sub
Private Sub t()
pt$ = "Tawny Mouse" & vbCrLf & vbCrLf & "A shy and quiet animology, but when you're by yourself, you are very active and peppy. Even if you are timid, you have a strong inner self." & vbCrLf & "Soul Mate: Gray Chicken" & vbCrLf & "In conflict with: Pink Sloth"
MsgBox pt$
End Sub
Private Sub u()
pu$ = "Bronze Goat" & vbCrLf & vbCrLf & "You are hard-working, but only think of yourself. Yet, you are very clever and often self-confident." & vbCrLf & "Soul Mate: Ocre Dolphin" & vbCrLf & "In conflict with: Beige Raccoon"
MsgBox pu$
End Sub
Private Sub v()
PV$ = "Saffron Rabbit" & vbCrLf & vbCrLf & "You are shy but pure, and quick. You are sensitive but friendly. However, you can never seem to stay in one place." & vbCrLf & "Soul Mate: Scarlet Bear" & vbCrLf & "In conflict with: Lavender Chhetah"
MsgBox PV$
End Sub
Private Sub w()
pw$ = "Fuchsia Possum" & vbCrLf & vbCrLf & "You are extremely talkative, peppy, and like to make a lot of noise. People sometimes find you annoying, but you don't care because you always love to have fun." & vbCrLf & "Soul Mate: Green Puppy" & vbCrLf & "In conflict with: Brown Iguana"
MsgBox pw$
End Sub
Private Sub x()
px$ = "Periwinkle Cow" & vbCrLf & vbCrLf & "You are lazy but content with your simple life. Aside from this, you are very helpful, even if you won't admit it." & vbCrLf & "Soul Mate: Orange Snake" & vbCrLf & "In conflict with: Scarlet Bear"
MsgBox px$
End Sub
Private Sub y()
py$ = "Lavender Cheetah" & vbCrLf & vbCrLf & "You are the fastest and most agile of all animologies. Because of this, sometimes you do things immediately without even stopping to think, and bad things occur to you." & vbCrLf & "Soul Mate: Indigo Beaver" & vbCrLf & "In conflict with: Saffron Rabbit"
MsgBox py$
End Sub
Private Sub z()
pz$ = "Tan Giraffe" & vbCrLf & vbCrLf & "You are much of a loner, but you are brave and are a quick thinker. You like helping others, but you keep it a secret." & vbCrLf & "Soul Mate: Red Jaguar" & vbCrLf & "In conflict with: Maroon Panda"
MsgBox pz$
End Sub
Private Sub ab()
pab$ = "Beige Raccoon" & vbCrLf & vbCrLf & "You are the nervous type, and are always unsure of yourself. Yet, you are nice, friendly and quite smart." & vbCrLf & "Soul Mate: Teal Cat" & vbCrLf & "In conflict with: Bronze Goat"
MsgBox pab$
End Sub
Private Sub ac()
pac$ = "Maroon Panda" & vbCrLf & vbCrLf & "You are very picky, but extremely exact and accurate about everything. It often takes you a long time to make decisions, but you often get good results." & vbCrLf & "Soul Mate: Black sheep" & vbCrLf & "In conflict with: Tan Giraffe"
MsgBox pac$
End Sub
Private Sub ad()
pad$ = "Ocre Dolphin" & vbCrLf & vbCrLf & "You are a good friend to others, and love to have fun. People lighten up because of you, even if sometimes you act pretty weird." & vbCrLf & "Soul Mate: Bronze goat" & vbCrLf & "In conflict with: Peach Pig"
MsgBox pad$
End Sub
Private Sub ae()
pae$ = "Magenta Squirrel" & vbCrLf & vbCrLf & "You are good company and are always in a good mood, but once in a while you become greedy and selfish." & vbCrLf & "Soul Mate: Peach Pig" & vbCrLf & "In conflict with: Black Sheep"
MsgBox pae$
End Sub
Private Sub af()
paf$ = "Silver and Red (SR) Wolf" & vbCrLf & vbCrLf & "You are strong and most of the time naive, but when it comes to love you know it all, you prefer to be quiet around Adults." & vbCrLf & "Soul Mate: Gold Falcon" & vbCrLf & "In conflict with: Maroon Panda"
MsgBox paf$
End Sub
Private Sub ag()
pag$ = "Gold Falcon" & vbCrLf & vbCrLf & "You are a very smart person and you are aggresive too, although you let your pride get in the way sometimes, you are very good in romantic situations." & vbCrLf & "Soul Mate: Silver and Red Wolf" & vbCrLf & "In conflict with: Teal Cat"
MsgBox pag$
End Sub
