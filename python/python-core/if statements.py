### -------------------- ###
### PYTHON IF STATEMENTS ###
### -------------------- ###

# BASIC IF-THEN STATEMENT
MyName = "Michelangelo"

if MyName == "Michelangelo":
   print("Hi there, Michelangelo.")

# BASIC IF-ELIF STATEMENT
MyLastName = "Leonardo"

if MyLastName == "Michelangelo":   
   print("Hi there, Michelangelo.")

elif MyLastName == "Leonardo":
   print("Hi there, Leonardo.")

# MULTI IF-ELIF STATEMENT
MySecLastName = "Raphael"

if MySecLastName == "Michelangelo":   
   print("Hi there, Michelangelo.")

elif MySecLastName == "Leonardo":
   print("Hi there, Leonardo.")

elif MySecLastName == "Raphael":
   print("Hi there, Raphael.")

# BASIC IF-ELIF-ELSE STATEMENT
MyMiddleName = "Donatello"

if MyMiddleName == "Michelangelo":   
   print("Hi there, Michelangelo.")

elif MyMiddleName == "Leonardo":
   print("Hi there, Leonardo.")

elif MyMiddleName == "Raphael":
   print("Hi there, Raphael.")

else:
   print("This must be Donatello! Hi there, Donatello.")

### --------------------------- ###
### PYTHON NESTED IF STATEMENTS ###
### --------------------------- ###

MyAge = 300
MyFavoriteColor = "Purple"

if MyAge >= 800:   
   
   if MyFavoriteColor == "Green":
      print("If so powerful you are, why leave?")

   elif MyFavoriteColor == "Red":
      print("I have waited a very long time for this, my little green friend.")

elif MyAge >= 300 or MyAge <= 799:

   if MyFavoriteColor == "Purple":
      print("This party's over.")

   elif MyFavoriteColor == "Red":
      print("You're impossibly outnumbered.")

else:
   print("I find your lack of coding disturbing.")
