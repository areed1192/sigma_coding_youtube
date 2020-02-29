Sub InternetExplorerObject()

Dim IEObject As InternetExplorer

'Create a new instance of the Internet Explorer Object
Set IEObject = New InternetExplorer

    'Make sure the app is visible
    IEObject.Visible = True
    
    'Navigate to a URL we specify.
    IEObject.Navigate Url:="https://youtube.com", Flags:=navOpenInNewWindow
    
    'One of the things we need to understand is that loading the page can take a while.
    'We should always wait for the page to load before continuing on to the next step.
    'This Loop will keep us waiting as long as the IEObject is in a Busy state or
    'the ReadyState does not communicate complete.
    Do While IEObject.Busy = True Or IEObject.ReadyState <> READYSTATE_COMPLETE
       
       'Wait one second, and then try again
       Application.Wait Now + TimeValue("00:00:01")
       
    Loop
    
    'Print the URL we are currently at.
    Debug.Print IEObject.LocationURL
    
    'Get the HTML document for the page
    Dim IEDocument As HTMLDocument
    Set IEDocument = IEObject.Document
    
    'The simplest thing we can do is grab an element, and then the inner text.
    Debug.Print IEDocument.getElementById("header").innerText
    
    'Grab a elements collection
    Dim IEElements As IHTMLElementCollection
    Set IEElements = IEDocument.getElementsByClassName("yt-shelf-grid-item yt-uix-shelfslider-item")
        Debug.Print IEElements.Length
        
    'Grab a specific element.
    Dim IEElement As IHTMLElement
    Set IEElement = IEElements.Item(2)
    
        'WARNING: THERE IS NO GUARANTEE THAT ALL THE ELEMENTS WILL CONTAIN THE SAME INFO
        Debug.Print "------------------"
        
        'Print the title
        Debug.Print "TITLE"
        Debug.Print IEElement.Title
        Debug.Print "------------------"
        
        'Print the className
        Debug.Print "CLASS NAME"
        Debug.Print IEElement.className
        Debug.Print "------------------"
        
        'Print the InnerHTML
        Debug.Print "INNER HTML"
        Debug.Print IEElement.innerHTML
        Debug.Print "------------------"
        
        'Print the ID
        Debug.Print "ID"
        Debug.Print IEElement.ID
        Debug.Print "------------------"
        
        'Print the Inner Text
        Debug.Print "INNER TEXT"
        Debug.Print IEElement.innerText
        Debug.Print "------------------"
    
        'Print the Parent Element
        Debug.Print "PARENT ELEMENT"
        Debug.Print IEElement.parentElement
        Debug.Print "------------------"
    
    'Grab all the anchors
    Dim IEAnchors As IHTMLElementCollection
    Set IEAnchors = IEDocument.anchors
    
    'Grab a specific anchor
    Dim IEAnchor As IHTMLAnchorElement
    Set IEAnchor = IEAnchors.Item(1)
        
        'Lets get some details about one of those links.
        Debug.Print IEAnchor.host
        Debug.Print IEAnchor.hostname
        Debug.Print IEAnchor.pathname
        Debug.Print IEAnchor.protocol
        Debug.Print IEAnchor.protocolLong
        Debug.Print IEAnchor.href
        Debug.Print "------------------"
       
    'Grab all the links in the document, these are basically the anchors
    Dim IELinks As IHTMLElementCollection
    Dim IELink As Object
    Set IELinks = IEDocument.Links
    Set IELink = IELinks.Item(1)
    
        Debug.Print IELink.href
        Debug.Print "------------------"
    
    'Grab all the style sheets in the document
    Dim IEStyleSheets As IHTMLStyleSheetsCollection
    Set IEStyleSheets = IEDocument.styleSheets
    
    'Grab a specific style sheet
    Dim IEStyleSheet As IHTMLStyleSheet
    Set IEStyleSheet = IEStyleSheets.Item(1)
        
        'Grab the CSS Text
        Debug.Print IEStyleSheet.cssText
        Debug.Print "------------------"
    
    'Grab all the images
    Dim IEImages As IHTMLElementCollection
    Dim IEImage As IHTMLImgElement
    
    'Grab a single image
    Set IEImages = IEDocument.images
    Set IEImage = IEImages.Item(1)
        
        'Get some information about our image
        Debug.Print IEImage.src
        Debug.Print IEImage.fileCreatedDate
        Debug.Print IEImage.fileSize
        Debug.Print IEImage.Height
        Debug.Print "------------------"
        
    Count = 1
    'Drop each link into an excel range
    For Each IEImage In IEImages
        Cells(Count, 1).Value = IEImage.src
        Count = Count + 1
    Next
    
    
End Sub
