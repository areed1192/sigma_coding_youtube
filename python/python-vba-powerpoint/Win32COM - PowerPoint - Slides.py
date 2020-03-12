import win32com.client

# Get the Active Instance of PowerPoint
PPTApp = win32com.client.GetActiveObject("PowerPoint.Application")

# Get the Active Presentation
PPTPresentation = PPTApp.ActivePresentation

# Let's see how many slides there are in our presentation
print("There are {} slides in my presentation".format(PPTPresentation.Slides.Count))

# With the slides collection we can actually grab a `range` of slides. That means we can specify multiple slides we want to "modify" at once.
PPTSldRng = PPTPresentation.Slides.Range([1, 4, 6])

# I want these slides to all have the same tags, so I'll set the entire SlideRange's Tag Property.
print("My Slide Range has {} slides in it.".format(PPTSldRng.Count))
print('')

# I can select multiple slides, using the select method.
PPTSldRng.Select()

# Loop through each slide in the Presentation, using the Slides Collection. Slides Collection returns all the Slide Ojbects.
for PPTSld in PPTPresentation.Slides:   

    # let's get some info on each slide
    print("Slide has an index of {}".format(PPTSld.SlideIndex))
    print("Slide has a Shape Count of {}".format(PPTSld.Shapes.Count))
    print("Slide has a Name of {}".format(PPTSld.Name))
    print("Slide has a Layout of {}".format(PPTSld.Layout))
    print("Slide has a Hyperlink count of {}".format(PPTSld.Hyperlinks.Count))
    print("Slide has a Comments count of {}".format(PPTSld.Comments.Count))
    print("Slide has a Tags count of {}".format(PPTSld.Tags.Count))
    print('')

    # I want to see if any of my slides have a 3D Model shape on it. mso3DModel
    for PPTShape in PPTSld.Shapes:
        
        # if the shape type matches
        if PPTShape.Type == 30:

            # Let's also add some tags to the slide
            for tag in [('Slide Type','Slide 3D'),('Slide Object','T-Rex'),('Object Status','Rotated')]:

                PPTSld.Tags.Add(Name = tag[0], Value = tag[1])
            
            print('This Slide has {} tags'.format(PPTSld.Tags.Count))
            
            # Grab the Model3D Property in it.
            Shape3D = PPTShape.Model3D

            # Print it out
            print("The Field of View is {}".format(Shape3D.FieldOfView))

            # Increment the rotation along the x-axis by the specified degress 0 to 360
            Shape3D.IncrementRotationX(50)

            for _ in range(1, PPTSld.Tags.Count):
                print(PPTSld.Tags.Name(_))        

            # I forgot!!! I need to add a comment that contains the tags!
            # first I need to grab all the tags again and put them in a list where one value is the name and one is the value.
            tags_list = [str((PPTSld.Tags.Name(_), PPTSld.Tags.Value(_)))for _ in range(1, PPTSld.Tags.Count)]
            
            # Then add a new comment to the slide which desginates me as the author and adds the tags.
            PPTComment = PPTSld.Comments.Add(Left = 12, Top = 12, Author = 'Alex Reed', AuthorInitials = 'AR', Text = ','.join(tags_list))



    