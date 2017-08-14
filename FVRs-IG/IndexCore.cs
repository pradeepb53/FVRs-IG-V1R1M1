using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text.RegularExpressions;


namespace FVRs_IG
{
    class IndexCore
    {
        //-------------------------------------------------------------------------------------------------------------//
        //                           Source modifications
        //-------------------------------------------------------------------------------------------------------------//
        // V1R0M0 -> V1R0M1 (Bug fixes)
        // 1) 04/24/2017 - On the printed index, 2 columns were skiped over to next page
        //     Reason    - \f (Form-feed character) was in the middle of the word("Self-represented\fSTEVEN" as one word)  
        //     Fix       - Add \f removal on to processTranscript() method
        //-------------------------------------------------------------------------------------------------------------//
        // V1R1M0 -> V1R1M0 (New version/release candidate)
        // 1) 08/14/2017 - Excluded word list is now saved and retrieved from a text document.Replaced current single word -
        //                 'Add Process', with a multi line text box and a 'Save' button to Add-Modify excluded words. 
        //-------------------------------------------------------------------------------------------------------------//


        private Application app = new Application();

        private List<TranscriptWord> WordIndexDictionary = new List<TranscriptWord>();
        private static string[] finalDeDupedWordList = null;

        private static Stopwatch applicationTime = new Stopwatch();

        private string currentWord = "";

        public IndexCore()
        {

        }
        public void processTranscript(string fileName, string[] excludedWordList)
        {
            applicationTime.Start(); //Start stopwatch
            Document document = app.Documents.Open(fileName, ReadOnly: true);

            //Split words into an array

            string firstParseString = "";
            string secondParseString = "";
            string thirdParseString = ""; // new changes

            firstParseString = document.Content.Text;

            string[] firstParseWordList = null;
            string[] secondParseWordList = null;
            string[] thirdParseWordList = null;

            //First parse-split, remove line-carrage, tabs after questions, tabs after answers, paranthesis,tab after ?, tab after period etc..etc

            firstParseWordList = firstParseString.Replace("\r", " ").Replace("Q\t", " ").Replace("A\t", " ").Replace("(", " ")
                      .Replace(")", " ").Replace("?\t", "  ").Replace("—\t", "  ").Replace("?", "  ")
                      .Replace(".\t", "  ").Split(' ');

            //Read after first-parse-split, remove tabs with leading digits

            for (int i = 0; i < firstParseWordList.Length; i++)
            {
                if (firstParseWordList[i].Trim().Equals(""))
                {
                    continue;
                }
                else
                {
                    if (i == firstParseWordList.Length - 1)
                    {
                        string tempOutputOne = Regex.Replace(firstParseWordList[i], @"^\d+\t", " "); //Remove "digits and tab" at the begining of the word - OK 
                        string tempOutputTwo = Regex.Replace(tempOutputOne, @"\d+\t", " "); //Remove "digits and tab" from anywhere in the word - OK     
                        secondParseString += tempOutputTwo;
                    }
                    else
                    {
                        string tempOutputOne = Regex.Replace(firstParseWordList[i], @"^\d+\t", " "); //Remove "digits and tab" at the begining of the word - OK 
                        string tempOutputTwo = Regex.Replace(tempOutputOne, @"\d+\t", " "); //Remove "digits and tab" from anywhere in the word - OK       
                        secondParseString += tempOutputTwo + " ";
                    }
                }
            }

            //Second parse-split, remove all remaining tabs and under-scores
            //04/24/2017 - V1R0M1 - Remove \f Form-feed character

            //secondParseWordList = secondParseString.Replace("\t", " ").Replace("__", " ").Split(' '); - V1R1M0 04/24/2017

            //Second parse-split, remove all remaining tabs and under-scores
            //04/24/2017 - V1R0M1 - Remove \f (Form-feed character)

            secondParseWordList = secondParseString.Replace("\t", " ").Replace("\f", " ").Replace("__", " ").Split(' ');  // 04/24/2017 - V1R0M1

            //Read after second-parse-split, remove leading and trailling hyphens, remove leading and trailling dashes
            //Remove trailling periods, commas, colons and semicolons

            for (int i = 0; i < secondParseWordList.Length; i++)
            {
                if (secondParseWordList[i].Trim().Equals(""))
                {
                    continue;
                }
                else
                {
                    if (i == secondParseWordList.Length - 1)
                    {
                        string tempOutputOne = Regex.Replace(secondParseWordList[i], @"^\-+|\-+$", " "); //Remove leading and trailling hyphens
                        string tempOutputTwo = Regex.Replace(tempOutputOne, @"^\—+|\—+$", " "); //Remove leading and trailling dashes
                        string tempOutputThree = Regex.Replace(tempOutputTwo, @"\.+$|\,+$|\:+$|\;+$", " "); //Remove trailling periods, commas, colons, semicolons

                        thirdParseString += tempOutputThree;
                    }
                    else
                    {

                        string tempOutputOne = Regex.Replace(secondParseWordList[i], @"^\-+|\-+$", " ");
                        string tempOutputTwo = Regex.Replace(tempOutputOne, @"^\—+|\—+$", " ");
                        string tempOutputThree = Regex.Replace(tempOutputTwo, @"\.+$|\,+$|\:+$|\;+$", " ");

                        thirdParseString += tempOutputThree + " ";
                    }
                }

            }

            thirdParseWordList = thirdParseString.Split(' ');

            finalDeDupedWordList = thirdParseWordList.Distinct().ToArray();

            document.Close();

            this.searchDocumentAndCreateWordDicionary(ref finalDeDupedWordList, fileName, excludedWordList);

        }

        private void searchDocumentAndCreateWordDicionary(ref string[] finalDeDupedWordList, string fileName, string[] excludedWordList)
        {
            Document document = app.Documents.Open(fileName, ReadOnly: true);
            document.Activate();

            HashSet<string> processedWordList = new HashSet<string>();

            string finalSearchWord = "";

            //Set console window properties
            Console.Title = "- Index Generator Status -";
            Console.ForegroundColor = ConsoleColor.Green;


            for (int i = 0; i < finalDeDupedWordList.Length; i++)
            {

                finalSearchWord = finalDeDupedWordList[i].Trim();

                //Words and sentences with double quotes(" ") should be identified, quotes should be removed in order to preserve correct print order (i.e. #'s $'s digits and actual words)  

                if (Regex.IsMatch(finalSearchWord, @"^[a-zA-Z0-9\$#]"))
                {

                }
                else
                {
                    //If the word is not all spaces and does not starts with one of the allowed charactors, then remove first position, could be a starting double quote or single quote 
                    if (finalSearchWord != "")
                    {
                        finalSearchWord = finalSearchWord.Remove(0, 1);
                    }

                }

                //If the last position of the word in not one of the allowed charactors, then remove it, could be a closing double quote or single quote 
                if (finalSearchWord != "")
                {
                    if (Regex.IsMatch(finalSearchWord.Substring(finalSearchWord.Length - 1, 1), @"[a-zA-Z0-9\$#]"))
                    {

                    }
                    else
                    {

                        finalSearchWord = finalSearchWord.Remove(finalSearchWord.Length - 1, 1);

                    }
                }

                //Cleanup any spaces created by above process, if any 
                finalSearchWord.Trim();

                if (finalSearchWord.Length > 2)
                {
                    //Check whether current word is already processed?
                    string processedWord = processedWordList.FirstOrDefault(w => w == finalSearchWord);

                    //Check whether word is in the excluded list?
                    int wordInExcludedList = Array.IndexOf(excludedWordList, finalSearchWord);

                    if ((processedWord == null) && (wordInExcludedList < 0)) //Not in the processed word list, not in the excluded list so it will be in the index
                    {

                        Console.WriteLine("Scanning transcript- processing word # " + i);

                        var CustomWord = new TranscriptWord();

                        // Not in the array, it is a new word so add to processed word list and start processing....

                        processedWordList.Add(finalSearchWord);

                        CustomWord.Name = finalSearchWord;

                        int wordFoundFrequency = 0;

                        Range searchRange = document.Range(Start: document.Content.Start, End: document.Content.End); //Look for the word from start of the transcript to end

                        searchRange.Find.Forward = true;
                        searchRange.Find.MatchCase = true;
                        searchRange.Find.Text = finalSearchWord;

                        currentWord = finalSearchWord;

                        searchRange.Find.Execute(MatchWholeWord: true);
                        int currentLineNumber = 0;
                        int currentPageNumber = 0;
                        int pageNumberOfTheWord = 0;
                        int lineNumberOfTheWord = 0;

                        string textOfTheSearchedRangeSentence = ""; //04/21/2017
                        string firstWordOfTheSearchedSentence = "";  //04/21/2017

                        while (searchRange.Find.Found)
                        {

                            Console.WriteLine("Looking for word : " + currentWord);

                            // If final search word is only a number, get current sentence being searched and extract the first word

                            if (Regex.IsMatch(finalSearchWord, @"^[0-9]"))   //04/21/2017
                            {
                                 textOfTheSearchedRangeSentence = searchRange.Sentences.First.Text;

                                if (textOfTheSearchedRangeSentence.Length >= finalSearchWord.Length) //04/21/2017 // To avoid "System.ArgumentOutOfRangeException"
                                {
                                    firstWordOfTheSearchedSentence = textOfTheSearchedRangeSentence.Substring(0, finalSearchWord.Length);
                                }    

                            }

                            //If sentence starts with a number, it usually is a question number, now if it is a number and matches the searched text, it definitely 
                            // cannot be a regular word, it got to be a question number, so ignore! 

                            if ((firstWordOfTheSearchedSentence == finalSearchWord) && (Regex.IsMatch(firstWordOfTheSearchedSentence, @"^[0-9]")))
                            {

                            }
                            else
                            {
                                //Process all pages, including the cover page
                                wordFoundFrequency++;

                                currentPageNumber = searchRange.Information[WdInformation.wdActiveEndPageNumber];

                                currentLineNumber = searchRange.Information[WdInformation.wdFirstCharacterLineNumber];


                                //Check whether current word is repeating in the same page and line number, if not, create the "Occurrence" object. 

                                if (wordFoundFrequency > 1)
                                {
                                    if (pageNumberOfTheWord != currentPageNumber || lineNumberOfTheWord != currentLineNumber)
                                    {
                                        pageNumberOfTheWord = currentPageNumber;
                                        lineNumberOfTheWord = currentLineNumber;

                                        var CustomOccurrence = new Occurrence { CustomPageNumber = pageNumberOfTheWord, CustomLineNumber = lineNumberOfTheWord };
                                        CustomWord.PageAndLine.Add(CustomOccurrence);
                                    }

                                }
                                else
                                {
                                    pageNumberOfTheWord = currentPageNumber;
                                    lineNumberOfTheWord = currentLineNumber;

                                    var CustomOccurrence = new Occurrence { CustomPageNumber = pageNumberOfTheWord, CustomLineNumber = lineNumberOfTheWord };
                                    CustomWord.PageAndLine.Add(CustomOccurrence);
                                }


                            }

                            searchRange.Find.Execute(MatchWholeWord: true);

                        }

                        CustomWord.Frequency = wordFoundFrequency;
                        WordIndexDictionary.Add(CustomWord);
                    }

                }
            }

            document.Close();
        }

        public void printWordIndex()
        {
            List<TranscriptWord> IndexWordDirectory = WordIndexDictionary.OrderBy(o => o.Name).ToList();

            Document indexDoc = app.Documents.Add();
            Range indexRange = indexDoc.Range();

            //Set columns 
            indexDoc.PageSetup.TextColumns.SetCount(5);

            indexDoc.Activate();
            indexRange.Select();

            //Temp vars
            string indexAlphabetLabel = "";

            //Logical vars
            bool isRealNumber = false;
            bool isNumberSign = false;
            bool isLetter = false;
            bool isCurrency = false;
            bool numberSignPrinted = false;
            bool numericalDigitsPrinted = false;
            bool currencySignPrinted = false;
            bool firstLetterPrinted = false;

            foreach (TranscriptWord item in IndexWordDirectory)
            {

                Console.WriteLine("Assembling index - processing word : " + item.Name);

                //Check first charactor of the word to determine type and identification
                string firstCharOfTheWord = item.Name.Substring(0, 1);

                //Assert whether first char of the current word is a number sign,$ sign,digit or letter 
                //then print it under appropriate label

                //Currency
                if (Regex.IsMatch(firstCharOfTheWord, @"^[$]"))
                {
                    isCurrency = true;
                    isLetter = false;
                    isRealNumber = false;
                    isNumberSign = false;
                }

                //Numerical Digit
                if (Regex.IsMatch(firstCharOfTheWord, @"^[0-9]"))
                {
                    isRealNumber = true;
                    isCurrency = false;
                    isLetter = false;
                    isNumberSign = false;
                }

                //Starts with number sign
                if (Regex.IsMatch(firstCharOfTheWord, @"^[#]"))
                {
                    isNumberSign = true;
                    isCurrency = false;
                    isLetter = false;
                    isRealNumber = false;
                }

                // Alphabetical letter
                if (Regex.IsMatch(firstCharOfTheWord, @"^[a-zA-Z]"))
                {
                    isLetter = true;
                    isCurrency = false;
                    isRealNumber = false;
                    isNumberSign = false;

                    if (indexAlphabetLabel.ToLower() == firstCharOfTheWord.ToLower())
                    {
                        firstLetterPrinted = true;
                    }
                    else
                    {
                        firstLetterPrinted = false;
                    }

                }

                // First char is $ sign, dollar sign is not printed.
                if ((isCurrency) && (!currencySignPrinted))
                {
                    indexAlphabetLabel = "$";

                    Paragraph labelParagraph = indexDoc.Paragraphs.Add();
                    labelParagraph.Range.Font.Size = 15;
                    labelParagraph.Range.Font.Bold = 1;
                    labelParagraph.Range.Text = " " + indexAlphabetLabel + " " + "\r\n";
                    currencySignPrinted = true;
                }

                // Real digits, label 0-9 not printed.
                if ((isRealNumber) && (!numericalDigitsPrinted))
                {
                    indexAlphabetLabel = "0-9";

                    Paragraph labelParagraph = indexDoc.Paragraphs.Add();
                    labelParagraph.Range.Font.Size = 15;
                    labelParagraph.Range.Font.Bold = 1;
                    labelParagraph.Range.Text = " " + indexAlphabetLabel + " " + "\r\n";
                    numericalDigitsPrinted = true;
                }

                // First char is number sign, number sign is not printed.
                if ((isNumberSign) && (!numberSignPrinted))
                {
                    indexAlphabetLabel = "#";

                    Paragraph labelParagraph = indexDoc.Paragraphs.Add();
                    labelParagraph.Range.Font.Size = 15;
                    labelParagraph.Range.Font.Bold = 1;
                    labelParagraph.Range.Text = " " + indexAlphabetLabel + " " + "\r\n";
                    numberSignPrinted = true;
                }

                //First char is a letter, alphabet label is not printed. 
                if ((isLetter) && (!firstLetterPrinted))
                {
                    indexAlphabetLabel = firstCharOfTheWord;

                    Paragraph labelParagraph = indexDoc.Paragraphs.Add();
                    labelParagraph.Range.ParagraphFormat.SpaceBefore = 0;
                    labelParagraph.Range.Font.Size = 15;
                    labelParagraph.Range.Font.Bold = 1;
                    labelParagraph.Range.Text = "- " + indexAlphabetLabel.ToUpper() + " -" + "\r\n";
                    firstLetterPrinted = true;
                }

                Paragraph currentWordParagraph = indexDoc.Paragraphs.Add();
                currentWordParagraph.Range.Font.Size = 10;
                currentWordParagraph.Range.Font.Bold = 1;

                currentWordParagraph.Range.Text = item.Name + " [" + item.Frequency + "]" + "\r\n";

                int columnsPerRowCount = 0;
                int columnOnePageNumber = 0;
                int columnOneLineNumber = 0;

                foreach (Occurrence step in item.PageAndLine)
                {
                    Console.WriteLine("Checking frequency of : " + item.Name);

                    columnsPerRowCount++;

                    if (columnsPerRowCount == 2)
                    {
                        Paragraph pageAndLineParagraph = indexDoc.Paragraphs.Add();
                        pageAndLineParagraph.Range.Font.Size = 7;
                        pageAndLineParagraph.Range.Font.Bold = 0;

                        pageAndLineParagraph.Range.Text = "[P" + columnOnePageNumber + ":" + "L" + columnOneLineNumber + "] [P" + step.CustomPageNumber + ":" + "L" + step.CustomLineNumber + "]" + "\r\n";
                        pageAndLineParagraph.Range.ParagraphFormat.SpaceAfter = 0;

                        columnsPerRowCount = 0;
                        columnOnePageNumber = 0;
                        columnOneLineNumber = 0;
                    }
                    else
                    {
                        columnOnePageNumber = step.CustomPageNumber;
                        columnOneLineNumber = step.CustomLineNumber;

                    }

                }

                //If columnsPerRowCount is 1, then print 1 line and reset the counter 
                if (columnsPerRowCount == 1)
                {
                    Paragraph para3 = indexDoc.Paragraphs.Add();
                    para3.Range.Font.Size = 7;
                    para3.Range.Font.Bold = 0;
                    para3.Range.Text = "[P" + columnOnePageNumber + ":" + "L" + columnOneLineNumber + "]" + "\r\n";
                    para3.Range.ParagraphFormat.SpaceAfter = 0;

                    columnsPerRowCount = 0;
                    columnOnePageNumber = 0;
                    columnOneLineNumber = 0;
                }

            }

            applicationTime.Stop();
            TimeSpan elapsedTime = applicationTime.Elapsed;

            Console.WriteLine("Index generation completed!- Duration - {0} hour(s):{1} minute(s):{2} second(s)",
                elapsedTime.Hours, elapsedTime.Minutes, elapsedTime.Seconds);

            try
            {
                indexDoc.Save();
                indexDoc.Close();
                app.Quit();
            }
            catch (System.Runtime.InteropServices.COMException e)
            {
                Console.WriteLine("Error occured, please contact IT bitch!! : " + e);

            }
            finally
            {
                app.Quit();
            }
        }
    }
}
