Imports System.IO
Imports System.Xml

Module Module1

    Sub parseXML()
        Dim m_xmlr As XmlTextReader
        'Create the XML Reader
        m_xmlr = New XmlTextReader("C:\temp\tmp\gadget.gmanifest")
        'Disable whitespace so that you don't have to read over whitespaces
        m_xmlr.WhitespaceHandling = WhitespaceHandling.None
        'read the xml declaration and advance to family tag
        m_xmlr.Read()
        'read the family tag
        m_xmlr.Read()
        'Load the Loop
        While Not m_xmlr.EOF
            'Go to the name tag
            m_xmlr.Read()
            'if not start element exit while loop
            If Not m_xmlr.IsStartElement() Then
                Exit While
            End If
            'Get the Gender Attribute Value
            Dim genderAttribute = m_xmlr.GetAttribute("gadget")
            'Read elements firstname and lastname
            m_xmlr.Read()
            m_xmlr.GetAttribute("about")
            m_xmlr.Read()
            'Get the firstName Element Value
            'Dim description = m_xmlr.ReadElementString("description")
            'Get the lastName Element Value
            'Dim lastNameValue = m_xmlr.ReadElementString("lastname")
            'Write Result to the Console
            'Console.WriteLine("Gender: " & genderAttribute _
            '& " FirstName: " & firstNameValue & " LastName: " _
            '& lastNameValue)
            Console.Write(genderAttribute)
        End While
        'close the reader
        m_xmlr.Close()
    End Sub
End Module
