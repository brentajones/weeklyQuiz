import os
import sys
import re
import datetime
import webbrowser

from docx import opendocx, getdocumenttext

s = webbrowser.get('safari')


# parse word document, code from docx documentation

if __name__ == '__main__':
    try:
        document = opendocx(sys.argv[1])
    except:
        print("Nope.")
        exit()

    sys.stderr.write("\x1b[2J\x1b[H")
    print "Getting document text"
    paratextlist = getdocumenttext(document)
    
    newparatextlist = []
    for paratext in paratextlist:
        newparatextlist.append(paratext.encode("utf-8"))
        
    numberedlist = enumerate(newparatextlist)
        
# Create empty lists of quiz components
    
    print "Analyzing document"
    questions = []
    answers = []
    correctans = []
    text = []
    url = []
    
    for i,item in numberedlist:
        # find lines ending with question marks
        m = re.search('.*\?$', item)
        if m:
            # add lines ending with question marks to questions lists
            questions.append(m.group(0))
            # create temporary list for answers
            answerstemp = []
            for j in range(0,4):
                # add next four lines to temporary answers list
                answerstemp.append(numberedlist.next()[1])
            # add temporary answers list, as a list, to answers list
            answers.append(answerstemp)
            # add next line to correct answer list
            correctans.append(numberedlist.next()[1])
            # add next line to text list
            text.append(numberedlist.next()[1])
            # add next line to URL list
            url.append(numberedlist.next()[1])

    # now we need to get the index of the correct answer in the list of answers by comparing the correct answer to the list. First, we zip the two together
    getcorrectind = zip(answers,correctans)
    # now enumerate the zipped list and create a new empty list for the indexes
    getcorrectindlist = enumerate(getcorrectind)
    correctind = []
    
    for i in range(10):
        currquestion = getcorrectindlist.next()[1]
        answerlist = currquestion[0]
        correctanswer = currquestion[1]
        # for each of the 10 answers, do a regex for the text after 'Answer: '
        m = re.search('Answer: (.*)',correctanswer)
        # find the matching index in the answer list, and add it to the correctind list
        indexes = (i for i,val in enumerate(answerlist) if re.match(m.group(1), val))
        correctind.append(list(indexes))
    
    print "Preparing quiz.js"
    # create the merged quiz list
    quiz = zip(questions,answers,correctind,text,url)
    
    # get today's date
    d = datetime.date.today()
    
    # create string for js filename
    filename_js = "js/" + d.strftime("%Y%m%d") + ".js"
    
    #create string for html filename
    filename_html = "embeds/" + d.strftime("%Y%m%d") + ".html"
    
    # get friday's date
    while d.weekday() != 4:
        d += datetime.timedelta(1)
        
    # create date for use in header info
    date = d.strftime("%b. %d, %Y")

        
    # create opening text for quizJSON
    opening = 'var quizJSON = {\n\t'

    info = '"info": {\n\t\t"name":\t"St. Louis Public Radio News Quiz for ' + date + '",\n\t\t"main":\t"<p>How well were you paying attention to the news this week? Take our quiz and find out.</p>",\n\t\t"results": "<h5>Thanks for playing. Keep listening and reading, then come back and take next week\'s quiz.</h5>",\n\t\t"level1":  "Nice work.",\n\t\t"level2":  "Almost there.",\n\t\t"level3":  "Try again next week." // no comma here\n\t},'

    questionsmeta = '\n\t"questions": ['
    
    # create questions strings
    questionstring = ''
    for j,question in enumerate(quiz):
        answerstring = ''
        # get the correct answer's index, then write the correct string for each answer, subbing in true/false as appropriate
        correctans = question[2][0]
        for i in range (0,4):
            truefalse = str(correctans == i)
            answerstring += '\t\t\t\t{"option": "' + question[1][i] + '",\t"correct": ' + truefalse.lower() + '}'
            if i < 3:
                answerstring += ',\n'

        # construct the questionstring for each question, including the question itself, the answerstring, and the answer messages and URLs
        questionstring += '\n\t\t{\n\t\t\t"q": "' + question[0] + '",\n\t\t\t"a": [\n' + answerstring + '\n\t\t\t],\n\t\t\t"correct": "<p><span>That\'s right!</span></p> <p><a target=\'_blank\' href=\'' + question[4] + '\'>' + question[3] + '</a></p>",\n\t\t\t"incorrect": "<p><span>Incorrect.</span></p> <p><a target=\'_blank\' href=\'' + question[4] + '\'>' + question[3] + '</a></p>"\n\t\t}'
        if j != 9:
            questionstring += ','

    # create content for the whole file
    filestring = opening + info + questionsmeta + questionstring + '\n\t]\n};'
    
    # create the file, write the content, close the file
    print "Writing quiz javascript"
    target = open (filename_js, 'w')
    target.write(filestring)
    target.close()
    
    #prompt user to create the corepublisher file and bit.ly link
    raw_input("Press enter to continue.")
    sys.stderr.write("\x1b[2J\x1b[H")
    print "Now we need to create the Core Publisher post"
    print "I'll open a browser window for you."
    raw_input("Press enter to continue.")
    s.open_new("http://news.stlpublicradio.org/node/add/post")
    sys.stderr.write("\x1b[2J\x1b[H")
    print "Enter the headline:"
    print "News Quiz for " + date
    print "---"
    print "Enter the byline:"
    print "Dale Singer"
    print "---"
    print "Enter the category:"
    print "Other News"
    print "---"
    print "Enter the slug:"
    print "News Quiz"
    print "---"
    print "Skip the content for now"
    print "---"
    print "Enter the tag:"
    print "quiz"
    print "---"
    print "Enter any related content you want"
    print "---"
    raw_input("Press enter to continue")
    sys.stderr.write("\x1b[2J\x1b[H")
    print "Now press the \"Save\" button"
    print "then the \"Preview\" button"
    print "and copy the web address (highlight it and press command-C)"
    raw_input("Press enter to continue")
    sys.stderr.write("\x1b[2J\x1b[H")
    print "Now we need to get a bit.ly link for the address"
    print "I'll open up bit.ly for you"
    raw_input("Press enter to continue")
    s.open_new("http://bit.ly")
    sys.stderr.write("\x1b[2J\x1b[H")
    print "Paste the address of the Core Publisher post into the \"Shorten URL\" box on bit.ly"
    print "Then copy the short URL to the clipboard"
    raw_input("Press enter to continue")
    sys.stderr.write("\x1b[2J\x1b[H")
    bitlyurl = raw_input("Paste the URL here and press enter:")
    
    
    
    
    # create html file content
    
    print "Preparing quiz html embed"
    htmlstring = """<!DOCTYPE html>
    <html>
        <head>
                <meta content="text/html; charset=utf-8" http-equiv="content-type">
        	<meta name="twitter:card" content="summary_large_image">
        	<meta name ="twitter:site" content="@stlpublicradio">
        	<meta name="twitter:creator" content="@dalesinger">
        	<meta name="twitter:title" content="St. Louis Public Radio News Quiz">
        	<meta name="twitter:description" content="St. Louis Public Radio news quiz for """ + date + """">
        	<meta name="twitter:image:src" content=img/twittercard_image-01.png>
                <title>St. Louis Public Radio Weekly News Quiz</title>
                <link href="../css/slickQuiz.css" media="screen" rel="stylesheet" type="text/css">
                <link href="../css/bootstrap.css" rel="stylesheet">
                <link href="../css/theme.css" rel="stylesheet">
                <script src="../js/jquery.js"></script>
                <script src="../js/share.min.js"></script>
	
        </head>

        <body class="body-embed">
    	   <div class="container embed"> 

    	    <div id="slickQuiz">
            <h1 class="quizName"><!-- where the quiz name goes --></h1>

            <div class="quizArea">
                <div class="quizHeader">
                    <!-- where the quiz main copy goes -->

                    <a class="button startQuiz" href="#">Get Started!</a>
                </div>

                <!-- where the quiz gets built -->
            </div>

            <div class="quizResults">
    		<div class="row">
    			<div class="col-xs-8">
    			    <h3 class="quizScore">You Scored: <span><!-- where the quiz score goes --></span></h3>

    		            <h3 class="quizLevel"><strong>Ranking:</strong> <span><!-- where the quiz ranking level goes --></span></h3>
    			    </div>
    			    <div class="col-xs-4">
    				    <div class="sharebutton"></div>
    				    </div>
    				    </div>
    		            <div class="quizResultsCopy">
    		                <!-- where the quiz result copy goes -->
    		            </div>
    		        </div>
    		</div>
            </div>
    </div>
    </div>

	
    	<script type="text/javascript">
    	$(function () {
    	    $('#slickQuiz').slickQuiz();
    	});	
	
    	$(function() {
        	    $('.sharebutton').share({
        	    	url: '""" + bitlyurl + """',
    		text_font: false,
    		facebook: {
    			image: 'img/twittercard_image-01.png',
    			name: 'St. Louis Public Radio\\'s weekly news quiz'
    			},
		
        	    });
    	    });
	
    	</script>
            <script src="../js/slickQuiz.js"></script>
    	<script src="../""" + filename_js + """"</script>

    <script type="text/javascript">
    var gaJsHost = (("https:" == document.location.protocol) ? "https://ssl." : "http://www.");
    document.write(unescape("%3Cscript src='" + gaJsHost + "google-analytics.com/ga.js' type='text/javascript'%3E%3C/script%3E"));
    </script>
    <script type="text/javascript">
     var pageTrackerA = _gat._getTracker("UA-2139719-1");
     pageTrackerA._initData();
     pageTrackerA._trackPageview();

     var pageTrackerB = _gat._getTracker("UA-1741309-15");
     pageTrackerB._initData();
     pageTrackerB._trackPageview();
    </script>
        </body>
    </html>"""
    
    # create the html file, write the content, close the file
    print "Writing quiz html embed"
    target = open (filename_html, 'w')
    target.write(htmlstring)
    target.close()
    
    # open html file in web browser
    url = '/' + filename_html
    path = os.getcwd()
    htmlfilename = "file://" + path + url
    
    #prompt user for body height
    raw_input("Press enter to continue")
    sys.stderr.write("\x1b[2J\x1b[H")
    print "Now I'll open the quiz so you can check it for accuracy."
    print "Take the quiz and stop at the results page. Don't close the quiz."
    raw_input("Press enter to open the quiz and continue")
    s.open_new(htmlfilename)
    sys.stderr.write("\x1b[2J\x1b[H")
    print "After you've taken the quiz and arrived at the results page,"
    print "right click on the quiz and choose \"Inspect Element\"."
    print "Hover over the \"<body>\" tag (near the top)."
    print "Look for the size of the quiz in the browser window."
    print "It will say \"1324 x _something_\" (probably around 4000-5000)"
    quizheight = input("Enter that number and press enter: ")
    quizheight = quizheight + 30
    
    # create Core Publisher HTML temp file
    
    sys.stderr.write("\x1b[2J\x1b[H")
    print "Now I'll create the HTML to put in the Core Publisher post"
    raw_input("Press enter to continue")
    sys.stderr.write("\x1b[2J\x1b[H")
    print "Go back to the Core Publisher post in the browser"
    print "Click the button in the toolbar of the post content that says \"<source>\""
    print "Copy and paste the following text into the Core Publisher post:"
    print ""
    corepubstring = """[asset-images[{"caption": "", "fid": "44511", "style": "card_280", "uri": "public://201403/twittercard_image-01.png", "attribution": ""}]]<p>Were you paying attention to St. Louis Public Radio this week?</p><p>Take our quiz below, then come back next week and try again, or challenge your friends.</p><iframe height=\"""" + str(quizheight) + """\" overflow="hidden" src="http://stlpublicradio.github.io/weeklyQuiz/""" + filename_html + """"width="580"></iframe>"""
    print corepubstring
    print ""