Module Module1

    Sub Main()
        Console.BackgroundColor = ConsoleColor.DarkGreen
        Console.ForegroundColor = ConsoleColor.DarkRed
        Console.Clear()

        Dim days() As String = {"first", "second", "third", "fourth", "fith", "sixth", "seventh", "eighth", "ninth", "tenth", "eleventh", "twelth"}
        Dim presents() As String = {"A partridge in a pear tree", "Two turtle doves", "Three french hens", "Four calling birds", "Five gold rings", "Six geese a-laying", "Seven swans a-swimming", "Eight maids a-milking", "Nine ladies dancing", "Ten lords a-leaping", "Eleven pipers piping", "Twelve drummers drumming"}
        Dim dayLoop As Integer
        Dim presentLoop As Integer
        Dim presentCount As Integer = 0
        Dim tempSentence As String
        Dim SAPI
        SAPI = CreateObject("SAPI.spvoice")
        SAPI.rate = 1
        SAPI.Voice = SAPI.GetVoices.item(0)
        Console.Siz

        For dayLoop = 1 To 12
            presentCount = presentCount + 1

            tempSentence = "On the " + days(dayLoop - 1) + " day of christmas my true love gave to me "
            Console.WriteLine(tempSentence)
            SAPI.speak(tempSentence)

            If presentCount > 1 Then
                presents(0) = "And a partridge in a pear tree"
            End If

            For presentLoop = presentCount To 1 Step -1
                tempSentence = presents(presentLoop - 1)
                Console.WriteLine(tempSentence)
                If presentLoop = 1 Then
                    SAPI.Voice = SAPI.GetVoices.item(2)
                    SAPI.rate = -8
                End If
                SAPI.speak(tempSentence)
                SAPI.rate = 1
                SAPI.Voice = SAPI.GetVoices.item(0)
            Next

        Next
        Console.ReadLine()
    End Sub

End Module
