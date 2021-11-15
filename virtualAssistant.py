


# Installed the following 
import random
from selenium import webdriver # pip install selenium
import win32com.client as wincl # pip install pypiwin32
import speech_recognition as sr # pip install SpeechRecognition # pip install pipwin # pipwin install pyaudio
import os
import time



# Store speech recognition in variable r
r = sr.Recognizer()
# Adjust how loud you must speak into mic for speech recognition
voiceVolume = 100
# Amount of time a user can wait before responding
voiceDelay = 50.0

# Path to chromedriver
PATH = "C:\Program Files (x86)\chromedriver_win32\chromedriver.exe"

# Computer voice
speak = wincl.Dispatch("SAPI.SpVoice")

# Computer name
compName = 'KB : '

# Computer greet responses
greetResponse = ["Hello There, What can I help you with?", "Hey, What would you like me to do?", "lets get started with one of the commands below", "I'm up and running so go ahead and tell me whats next"]

# Stop the conversation
stopConvo = "nevermind"

# Computer wishes farewell
compFarewell = ["Alright then, we can talk later", "Goodbye", "Talk to you later then"]



# Start listening for greeting
with sr.Microphone() as source:
	print("\n\n\n\tSay Hello : ")
	# Set dynamic energy threshold to False to cancel out background noise
	r.dynamic_energy_threshold = False
	# Set to adjust mic sensitivity to voice
	r.energy_threshold = voiceVolume
	# Listen an account for a possible delay in speech
	audio = r.listen(source, timeout = voiceDelay)
	
	# Beginning of exception handler
	try:
	
		# Recieve and display greeting
		text = r.recognize_google(audio)
		print("\n\tYou said : {}".format(text))
		
		# If word or phrase is in users greeting, do the following.
		if "hello" in text: 
			
			# Text gives further instructions & Compter voice will give random greet back
			print ("\n\t" + compName + "Please use one of the following commands below to get started.")
			speak.Speak(random.choice(greetResponse)) 
				
			
			while True: # Loop following code forever until exception handler throws exception
			
				# Choose yes to continue or no to stop the script
				choice = input("\n\tEnter y/n to continue : ")
				
				# If yes, do the follwing	
				if "y" in choice:
				
					# Display options
					print ("\n\n\n\tGo Straight To Site")
					print ("\t---------------------------")
					print ("\tOpen Google")
					print ("\tOpen Internet Explorer")
					print ("\t---------------------------")
					print ("\tOpen Microsoft Word")
					print ("\tOpen Microsoft PowerPoint")
					print ("\tOpen Microsoft Excel")
					print ("\t---------------------------")
					print ("\tOpen VirtualBox")
					print ("\t---------------------------")
					print ("\tOpen Adobe Creative Cloud")
					print ("\t---------------------------")
					print ("\tText To Speech")
					print ("\tSay Nevermind to stop")
					
					# Listen for the selected spoken option
					with sr.Microphone() as source:
						print("\n\n\n\tSpeak your selection : ")
						r.dynamic_energy_threshold = False
						r.energy_threshold = voiceVolume
						audio = r.listen(source, timeout = voiceDelay)
						
						# Recieve and display selected spoken option
						text = r.recognize_google(audio)
						print("\n\tYou said : {}".format(text))
						
						# If site is spoken, do the following
						if "site" in text:
							speak.Speak("Please enter the website name....")
							# Type the url of a website and store in URL variable
							URL = input("\n\t" + "Input website URL : ")
							# If user adds www before site name, do the following
							if "www" in URL:
								protocol = "https://"
								# Open chromedriver from file path
								driver = webdriver.Chrome(PATH)
								# Concatenate protocol variable with website entered 
								driver.get(protocol + URL)
								# Pause before any further suggestions
								time.sleep(5)
							else: # If www is not entered before site name
								protocol = "https://www."
								# Open chromedriver from file path
								driver = webdriver.Chrome(PATH)
								# Concatenate protocol variable with website entered
								driver.get(protocol + URL)
								# Pause before any further suggestions
								time.sleep(5)
								
						# If the word Google is spoken, do the following
						elif "Google" in text:
							os.startfile("C:\Program Files (x86)\Google\Chrome\Application\chrome.exe")
							# Compter alerts the user Google is opening
							print ("\n\t" + compName + "I am now opening Google Chrome....")
							speak.Speak("I am now opening Google Chrome....")
						
						# Else if Internet Explorer is spoken
						elif "Internet Explorer" in text:
							ie = wincl.Dispatch("InternetExplorer.Application") 
							ie.Visible = True
							# Compter alerts the user Internet Explorer is opening
							print ("\n\t" + compName + "I am now opening Internet Explorer....")
							speak.Speak("I am now opening Internet Explorer....")
							
						# Else if Microsoft Word is spoken	
						elif "Word" in text:
							word = wincl.Dispatch("Word.Application")
							word.Visible = True
							# Compter alerts the user Microsoft Word is opening
							print ("\n\t" + compName + "I am now opening Microsoft Word....")
							speak.Speak("I am now opening Microsoft Word....")
							
						# Else if Microsoft Powerpoint is spoken	
						elif "PowerPoint" in text:
							powerPoint = wincl.Dispatch("PowerPoint.Application")
							powerPoint.Visible = True
							# Compter alerts the user Microsoft Powerpoint is opening
							print ("\n\t" + compName + "I am now opening Microsoft PowerPoint....")
							speak.Speak("I am now opening Microsoft PowerPoint....")
							
						# Else if Microsoft Excel is spoken	
						elif "Excel" in text:
							excel = wincl.Dispatch("Excel.Application")
							excel.Visible = True
							# Compter alerts the user Microsoft Excel is opening
							print ("\n\t" + compName + "I am now opening Microsoft Excel....")
							speak.Speak("I am now opening Microsoft Excel....")
							
						# Else if VirtualBox is spoken	
						elif "virtualbox" in text:
							os.startfile("C:\Program Files\Oracle\VirtualBox\VirtualBox.exe")
							# Compter alerts the user VirtualBox is opening
							print ("\n\t" + compName + "I am now opening VirtualBox....")
							speak.Speak("I am now opening VirtualBox....")
							
						# Else if Adobe Creative Cloud is spoken	
						elif "Adobe" in text:
							os.startfile("C:\Program Files (x86)\Adobe\Adobe Creative Cloud\ACC\Creative Cloud.exe")
							# Compter alerts the user Adobe Creative Cloud is opening
							print ("\n\t" + compName + "I am now opening Adobe Creative Cloud....")
							speak.Speak("I am now opening Adobe Creative Cloud....")
							
						# Else if text to speech is spoken
						elif "text to speech" in text:
							speak.Speak("Enter some text for me to read")
							# Insert text to be read
							speakText = input("\n\t" + "Enter Some Text Here: ")
							# Computer speaks text
							speak.Speak(speakText) 
							
							while True: # Loop until broken
								speak.Speak("Would you like me to read the text again?")
								with sr.Microphone() as source:
									print("\n\tSpeak Yes or No : ")
									r.dynamic_energy_threshold = False
									r.energy_threshold = voiceVolume
									audio = r.listen(source, timeout = voiceDelay)
									
									text = r.recognize_google(audio)
									print("\n\tYou said : {}".format(text))
									
									# If yes is spoken, read text again
									if "yes" in text:
										speak.Speak(speakText)
										spokenText = speakText
										print("\n\t" + spokenText)
									
									# If no is spoken, break the loop and move on
									elif "no" in text:
										break
										
									# Keep asking to reread text till correct answer is given
									else:
										speak.Speak("Sorry but you have to answer with yes or no")
										continue
							
						# Else if Nevermind is spoken, nothing happens	
						elif ((text) == stopConvo):
							# Computer prints and says goodbye
							print ("\n\t" + compName + "Reload script to start over")
							speak.Speak(random.choice(compFarewell))
						
						# One of the options displayed has to be selected
						else:
							print ("\n\t" + compName + "You did not use one of the following commands above.")
							speak.Speak("Please use one of the following commands above.")
						
					continue
					
					
				# Else if no, computer says goodbye and stops the script
				elif ("n" in choice):
					speak.Speak(random.choice(compFarewell))
					break
				
				# If anything besides y or n is entered, do the following
				else:
					speak.Speak("You must enter y for yes or n for no so we can continue.")
		else:
			print("\n\t" + compName + "Waiting for you to say hello")
			speak.Speak("Waiting for a hello")
	# Ending of exception handler - Any of the spoken input was no good
	except:
		
		# Alert the user that the input was misunderstood
		print("\n\t" + compName + "Sorry I do not understand your vocal command")
		speak.Speak("Sorry I do not understand your vocal command")