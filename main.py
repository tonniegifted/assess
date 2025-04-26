from getpass import win_getpass
# from db import dbfunc
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.QtCore import *
from PyQt5.uic import loadUi
from PyQt5 import QtWidgets
import sys
import mysql.connector

# Step 1: Enable High DPI scaling (must be before app is created)
QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)

# Step 2: Connect to DB (unchanged)
db = mysql.connector.connect(
    host="localhost",
    user="root",
    password="print",
    database="remedial",
    buffered=True
)
cur = db.cursor()

# Step 3: Main window class
class mainwin(QMainWindow):
    def __init__(self):
        super().__init__()
        loadUi("home.ui", self)
        self.loadschoolname()
        self.loadterm()
        self.enterscoresbutton.clicked.connect(self.loadenterscore)
        self.exitbutton.clicked.connect(self.close_application)
        self.updatebutton.clicked.connect(self.loadupdatescore)
        self.deletebutton.clicked.connect(self.loaddeletescore)
        self.adminbutton.clicked.connect(self.loadadminpanel)
        

    def loadschoolname(self):
        path="D:/TONNIEGIFTED/Documents/programs/Remedial2/name.txt"
        with open(path,"r") as file:
            schoolname=file.read()
            self.schoolnamelabel.setText(schoolname)

    def loadterm(self):
        cur.execute("SELECT selected_term FROM term WHERE is_active=1")
        term=cur.fetchone()[0]
        self.termlabel.setText(term)

    #Loading windows/screens
    def loadenterscore(self):
        enterscorewin=enterscore()
        widget.addWidget(enterscorewin)
        widget.setCurrentIndex(widget.currentIndex()+1)

    def loadupdatescore(self):
        updatescorewin=updatescore()
        widget.addWidget(updatescorewin)
        widget.setCurrentIndex(widget.currentIndex()+1)
    
    def loaddeletescore(self):
        deletescorewin=deletescore()
        widget.addWidget(deletescorewin)
        widget.setCurrentIndex(widget.currentIndex()+1)
        
    def loadadminpanel(self):
        adminwin=adminpanel()
        widget.addWidget(adminwin)
        widget.setCurrentIndex(widget.currentIndex()+1)

    def close_application(self):
        """Exit confirmation with system bell and complete window cleanup"""
        # Play system alert sound
        QApplication.beep()

        # Show confirmation dialog
        reply = QMessageBox.question(
            self,
            'Confirm Exit',
            'Are you sure you want to\nexit the program?',
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )

        if reply == QMessageBox.Yes:
            # Close all application windows
            for window in QApplication.topLevelWidgets():
                window.close()

            # Ensure complete application termination
            QApplication.quit()


class enterscore(QMainWindow):
    def __init__(self):
        super().__init__()
        loadUi("enterscores.ui", self)

        # Initialize grade and subject
        self.update_grade_subject()
        # Load initial data
        self.hidetotalscore()
        self.displayentered()
        self.loadlearners(self.grade)

        # Connect signals
        self.gradecombo.currentTextChanged.connect(self.update_grade_subject)
        self.gradecombo.currentTextChanged.connect(self.loadlearners)
        self.gradecombo.currentTextChanged.connect(self.displayentered)
        self.totalbutton.clicked.connect(self.savetotalscore)
        self.subjectcombo.currentTextChanged.connect(self.update_grade_subject)
        self.subjectcombo.currentTextChanged.connect(self.hidetotalscore)
        self.enterscoresbutton.clicked.connect(self.savescores)
        self.homebutton.clicked.connect(self.homepage)
    
        
    def displayentered(self):
        try:
            self.enteredscores.clear()
            # Get grade_id from combobox
            cur.execute("""SELECT grade_id FROM grade WHERE grade_name = %s""",
                    (self.gradecombo.currentText(),)) 
            grade_id = cur.fetchone()[0]
            
            # Get active term_id
            cur.execute("""SELECT term_id FROM term WHERE is_active = 1""")
            term_id = cur.fetchone()[0]
            
            subject_ids = [(1,), (2,), (3,), (4,), (5,), (6,), (7,), (8,), (9,)]
            output_text = ""

            for subject_id in subject_ids:
                cur.execute("""
                    SELECT s.subject_score, a.subject_abbr FROM score s
                    JOIN subject a ON a.subject_id = s.subject_id
                    WHERE s.subject_id = %s AND s.grade_id = %s AND s.term_id = %s 
                    LIMIT 1
                """, (subject_id[0], grade_id, term_id))
                
                result = cur.fetchone()
                if result:
                    output_text += f"{result[1]} | "

            output_text = output_text.rstrip(" | ")
            self.enteredscores.setText(output_text)

            # Add the readyforanalysis check
            required_subjects = {"MATHS", "ENG", "KISW", "SST", "INT", "AGN", "CRE", "CAS", "PTC"}
            if all(subject in output_text for subject in required_subjects):
                self.enteredscores.clear()
                self.enteredscores.setText(f"Grade scores ready for analysis")
                self.enteredscores.setStyleSheet("color:red; font-size:9px")
        except Exception as e:
            QMessageBox.critical(self,"AssessmentBoy",f"{e}")

    def update_grade_subject(self):
        """Update grade and subject when combobox changes."""
        self.grade = self.gradecombo.currentText()
        self.subject = self.subjectcombo.currentText()

    def hidetotalscore(self):
        """Disable total score widgets if the subject's total already exists."""
        try:
            self.totalfield.clear()
            #checking whether total score has been filled before proceeding to
            #input scores
            cur.execute("""
                SELECT total_score FROM total 
                WHERE term_id = (SELECT term_id FROM term WHERE is_active = 1)
                AND grade_id = (SELECT grade_id FROM grade WHERE grade_name = %s)
                AND subject_id = (SELECT subject_id FROM subject WHERE subject_abbr = %s)
            """, (self.grade, self.subject))

            total = cur.fetchone()
            #checking where subject scores has already been filled
            cur.execute("""SELECT subject_score FROM score WHERE subject_id=(
                                                       SELECT subject_id FROM subject WHERE subject_abbr=%s)AND grade_id=
                                                       (SELECT grade_id FROM grade WHERE grade_name=%s)""",
                        (self.subject, self.grade))
            subject = cur.fetchall()

            if total:

                self.totalbutton.setDisabled(True)
                self.totalfield.setDisabled(True)
                self.totalfield.setText(str(total[0]))
                self.enterscoresbutton.setDisabled(False)


                if subject:
                    self.enterscoresbutton.setDisabled(True)
            else:
                self.totalbutton.setDisabled(False)
                self.totalfield.setDisabled(False)
                self.enterscoresbutton.setDisabled(True)
                if not subject:
                    self.enterscoresbutton.setDisabled(False)



        except Exception as e:

            QMessageBox.critical(self, "Error", "Failed to fetch total score.")

    def loadlearners(self, grade):
        """Loading learners detail from the database"""
        self.scorestable.setColumnWidth(0, 40)   # Column 0 (learner_id)
        self.scorestable.setColumnWidth(1, 250)  # Column 1 (name)
        self.scorestable.setColumnWidth(2, 40)  # Column 2 (score)

        cur.execute("""
            SELECT learner_id, first, second, surname
            FROM learner
            WHERE grade = %s
        LIMIT 5""", (self.grade,))
        
        result = cur.fetchall()
        
        self.scorestable.setRowCount(len(result))
        
        for row, (learner_id, first, second, surname) in enumerate(result):
            name = f"{first} {second} {surname}"
            
            self.scorestable.setItem(row, 0, QTableWidgetItem(str(learner_id)))  # ID
            self.scorestable.setItem(row, 1, QTableWidgetItem(name))             # Full name
            self.scorestable.setItem(row, 2, QTableWidgetItem(""))               # Score (empty for now)

        self.scorestable.verticalHeader().setDefaultSectionSize(15)  # Or 25, 35 etc
        self.scorestable.verticalHeader().setFixedWidth(40)  # Adjust the number as needed
        self.hidetotalscore()

    def savetotalscore(self):
        """Saving total score per learning area/subject and per grade using upsert"""
        # subject = self.subjectcombo.currentText()
        try:
            total=self.totalfield.text()
            if len(total)==0:
                QMessageBox.information(self,"AssessmentBoy","Total Field is Required")
                return
            else:
        
                total = int(self.totalfield.text())

            # First get all the required IDs
            cur.execute("""
                SELECT 
                    (SELECT term_id FROM term WHERE is_active = 1 LIMIT 1) as term_id,
                    (SELECT subject_id FROM subject WHERE subject_abbr = %s LIMIT 1) as subject_id,
                    (SELECT grade_id FROM grade WHERE grade_name = %s LIMIT 1) as grade_id
            """, (self.subject, self.grade))

            ids = cur.fetchone()

            if ids and all(ids):  # Check all IDs exist
                term_id, subject_id, grade_id = ids

                # Use INSERT ON DUPLICATE KEY UPDATE
                cur.execute("""
                    INSERT INTO total (term_id, grade_id, subject_id, total_score)
                    VALUES (%s, %s, %s, %s)
                    ON DUPLICATE KEY UPDATE total_score = VALUES(total_score)
                """, (term_id, grade_id, subject_id, total))

                db.commit()
                QMessageBox.information(self,"AssessmentBoy","Total score saved successfully")
                self.totalfield.clear()
                self.hidetotalscore()

            else:
                QMessageBox.critical(self,"AssessmentBoy","Could not find all required IDs")

        except Exception as e:
            db.rollback()
            QMessageBox.critical(self,"AssessmentBoy",f"Error saving total score: {e}")
    def savescores(self):
        # Get IDs
        cur.execute("SELECT grade_id FROM grade WHERE grade_name=%s", (self.grade,))
        grade_id = cur.fetchone()[0]
        
        cur.execute("SELECT subject_id FROM subject WHERE subject_abbr=%s", (self.subject,))
        subject_id = cur.fetchone()[0]
        
        cur.execute("SELECT term_id FROM term WHERE is_active=1")
        term_id = cur.fetchone()[0]

        # Validate total
        if not self.totalfield.text():
            QMessageBox.critical(self, "AssessmentBoy", "Please set the total score first")
            return
        total = int(self.totalfield.text())

        for row in range(self.scorestable.rowCount()):
            learner_id = int(self.scorestable.item(row, 0).text())
            score_text = self.scorestable.item(row, 2).text()
            
            if not score_text:
                QMessageBox.critical(self, "AssessmentBoy", f"Please enter score for row {row+1}")
                return
                
            score = int(score_text)
            
            if score > total:
                QMessageBox.critical(self, "AssessmentBoy", 
                                f"Score {score} exceeds total {total} in row {row+1}")
                return

            # Calculate percentage
            percentage_score = (score / total) * 100
            rounded_score = round(percentage_score)
            
            if percentage_score <= 29:
                grading_level = "BE"
            elif percentage_score <= 49:
                grading_level = "AE"
            elif percentage_score <= 99:
                grading_level = "ME"
            else:
                grading_level = "EE"

            try:
                # Save/update subject score
                cur.execute("""
                    INSERT INTO score(grade_id, learner_id, subject_id, term_id, 
                                    subject_score, expectation)
                    VALUES(%s, %s, %s, %s, %s, %s)
                    ON DUPLICATE KEY UPDATE
                    subject_score = VALUES(subject_score),
                    expectation = VALUES(expectation)
                """, (grade_id, learner_id, subject_id, term_id, 
                    rounded_score, grading_level))
                
                # Calculate FRESH grand total by summing all subjects for this term
                cur.execute("""
                    SELECT COALESCE(SUM(subject_score), 0) 
                    FROM score 
                    WHERE learner_id = %s 
                    AND term_id = %s 
                    AND grade_id = %s
                """, (learner_id, term_id, grade_id))
                new_grandtotal = cur.fetchone()[0]
                
                # Update grand total
                cur.execute("""
                    INSERT INTO grand(learner_id, term_id, grade_id, grandtotal)
                    VALUES(%s, %s, %s, %s)
                    ON DUPLICATE KEY UPDATE
                    grandtotal = VALUES(grandtotal)
                """, (learner_id, term_id, grade_id, new_grandtotal))
                
                db.commit()
                
            except Exception as e:
                db.rollback()
                QMessageBox.critical(self, "Database Error", 
                                f"Failed to save scores: {str(e)}")
                return

        # Clear fields
        for row in range(self.scorestable.rowCount()):
            self.scorestable.item(row, 2).setText("")
            
        self.enterstatusbar.showMessage("Scores saved successfully", 3000)
        self.hidetotalscore()
        self.displayentered()
        # dbfunc()
        
    def homepage(self):
        home=mainwin()
        widget.setCurrentIndex(widget.currentIndex()+1)
        widget.addWidget(home)

class updatescore(QMainWindow):
    def __init__(self):
        super().__init__()
        loadUi("updatescores.ui", self)
        self.totalbutton.clicked.connect(self.loadlearnerlist)
        # self.grade = self.gradecombo.currentText()
        # self.subject = self.subjectcombo.currentText()
        self.gradecombo.currentTextChanged.connect(self.loadlearnerlist)
        self.listcombo.currentTextChanged.connect(self.loadlearnerscore)
        self.updatescorebutton.clicked.connect(self.updatescore)
        self.homebutton.clicked.connect(self.tohome)

    def loadlearnerlist(self):
        grade = self.gradecombo.currentText()
        subject = self.subjectcombo.currentText()
        self.listcombo.clear()

        cur.execute("""SELECT learner_id,first,second, surname FROM learner
        WHERE grade=%s""",(grade,))
        l=cur.fetchall()
        for i in l:
            learner=f"{i[0]}. {i[1]} {i[2]}"

            self.listcombo.addItem(str(learner))

    def loadlearnerscore(self):
        self.scorefield.clear()
        #retrieving total score
        cur.execute("""SELECT total_score FROM total WHERE term_id=(SELECT term_id FROM
        term WHERE is_active=1) AND subject_id=(SELECT subject_id FROM subject WHERE subject_abbr=%s)
        AND grade_id=(SELECT grade_id FROM grade WHERE grade_name=%s)""",(self.subjectcombo.currentText(),
                                                                          self.gradecombo.currentText()))
        total=cur.fetchone()
        if total:
            total_score=total[0]
            total_score=int(total_score)
        selected = self.listcombo.currentText()
        learner_id = selected.split(".")[0].strip()
        # learner_id=int(learner_id)
        subject_abbr=self.subjectcombo.currentText()
        cur.execute("""SELECT subject_score FROM score WHERE subject_id=(
        SELECT subject_id FROM subject WHERE subject_abbr=%s
        AND learner_id=%s)""",(subject_abbr,learner_id))
       
        s=cur.fetchone()
        if s:
            score=s[0]
            raw_score=int(score)/100*int(total_score)
            raw_score=round(raw_score)
            self.scorefield.setText(str(raw_score))
            self.scorefield.setDisabled(False)
            self.updatescorebutton.setDisabled(False)
        else:
            self.scorefield.setDisabled(True)
            self.updatescorebutton.setDisabled(True)

    def updatescore(self):
        try:
            # Get learner ID from combo box
            learner_id = int(self.listcombo.currentText().split(".")[0])
            
            # Validate score input
            score = self.scorefield.text()  # Changed from undefined 'score' variable
            if not score:
                QMessageBox.information(self, "AssessmentBoy", "Score field is required")
                return
            
            score = int(score)
            
            # Get subject ID
            cur.execute("SELECT subject_id FROM subject WHERE subject_abbr=%s", 
                    (self.subjectcombo.currentText(),))
            subject_id = cur.fetchone()[0]
            
            # Get total score for comparison
            cur.execute("""
                SELECT total_score FROM total 
                WHERE term_id = (SELECT term_id FROM term WHERE is_active=1) 
                AND subject_id = %s
                AND grade_id = (SELECT grade_id FROM grade WHERE grade_name=%s)
            """, (subject_id, self.gradecombo.currentText()))
            
            total_result = cur.fetchone()
            if not total_result:
                QMessageBox.critical(self, "AssessmentBoy", "Total score not found for this subject")
                return
                
            total_score = int(total_result[0])
            
            # Calculate percentage score
            subject_score = round((score / total_score) * 100)
            if subject_score > 100:
                QMessageBox.critical(self, "AssessmentBoy", "Score cannot exceed total score")
                return
            
            # Update subject score
            cur.execute("""
                UPDATE score 
                SET subject_score = %s 
                WHERE learner_id = %s 
                AND subject_id = %s
                AND term_id = (SELECT term_id FROM term WHERE is_active=1)
            """, (subject_score, learner_id, subject_id))
            
            # Update grand total (properly calculated)
            cur.execute("""
                SELECT COALESCE(SUM(subject_score), 0) 
                FROM score 
                WHERE learner_id = %s
                AND term_id = (SELECT term_id FROM term WHERE is_active=1)
            """, (learner_id,))
            
            newgrandtotal = cur.fetchone()[0]
            
            cur.execute("""
                UPDATE grand 
                SET grandtotal = %s 
                WHERE learner_id = %s
                AND term_id = (SELECT term_id FROM term WHERE is_active=1)
            """, (newgrandtotal, learner_id))
            
            db.commit()
            self.scorefield.clear()
            self.updatestatusbar.showMessage("Score updated successfully", 3000)
            
        except ValueError:
            QMessageBox.critical(self, "AssessmentBoy", "Please enter valid numeric values")
        except Exception as e:
            db.rollback()
            QMessageBox.critical(self, "AssessmentBoy", f"Error updating score: {str(e)}")
    
    def tohome(self):
        screen=mainwin()
        widget.addWidget(screen)
        widget.addWidget(screen)
        widget.setCurrentIndex(widget.currentIndex()+1)
    
class deletescore(QMainWindow):
    def __init__(self):
        super().__init__()
        loadUi("deletescores.ui", self)
        
        self.subjectdelete.clicked.connect(self.deletesubject)
        self.homebutton.clicked.connect(self.tohome)
        self.gradedelete.clicked.connect(self.deletegradescore)
        self.subdelete.clicked.connect(self.deletelsubject)
        self.deleteall.clicked.connect(self.deletelall)
        
    def deletesubject(self):
        resp = QMessageBox.question(
            self,
            "AssessmentBoy",
            f"Are you sure you want to delete\n{self.subjectcombo.currentText()} for Grade {self.gradecombo.currentText()}?",
            QMessageBox.Yes | QMessageBox.No
        )
        
        if resp != QMessageBox.Yes:
            return  # Exit if user clicks "No"

        try:
            grade_name = self.gradecombo.currentText()
            subject_abbr = self.subjectcombo.currentText()

            # 1. FIRST, fetch the subject_score BEFORE deleting it
            cur.execute("""
                SELECT subject_score FROM score
                WHERE subject_id = (SELECT subject_id FROM subject WHERE subject_abbr = %s)
                AND grade_id = (SELECT grade_id FROM grade WHERE grade_name = %s)
                AND term_id = (SELECT term_id FROM term WHERE is_active = 1)
            """, (subject_abbr, grade_name))

            subject_score_result = cur.fetchone()
            if subject_score_result:  # Only proceed if the score exists
                subject_score = subject_score_result[0]

                # 2. Fetch the current grand total
                cur.execute("""
                    SELECT grandtotal FROM grand 
                    WHERE grade_id = (SELECT grade_id FROM grade WHERE grade_name = %s) 
                    AND term_id = (SELECT term_id FROM term WHERE is_active = 1)
                """, (grade_name,))
                
                grandtotal_result = cur.fetchone()
                if grandtotal_result:
                    grandtotal = grandtotal_result[0]
                    new_grand = grandtotal - subject_score
                    # 3. Update the grand total BEFORE deleting the score
                    cur.execute("""
                        UPDATE grand 
                        SET grandtotal = %s
                        WHERE grade_id = (SELECT grade_id FROM grade WHERE grade_name = %s)
                        AND term_id = (SELECT term_id FROM term WHERE is_active = 1)
                    """, (new_grand, grade_name))

            # 4. NOW delete the score and related records
            cur.execute("""
                DELETE FROM score 
                WHERE grade_id = (SELECT grade_id FROM grade WHERE grade_name = %s)
                AND subject_id = (SELECT subject_id FROM subject WHERE subject_abbr = %s)
                AND term_id = (SELECT term_id FROM term WHERE is_active = 1)
            """, (grade_name, subject_abbr))

            cur.execute("""
                DELETE FROM total 
                WHERE term_id = (SELECT term_id FROM term WHERE is_active = 1)
                AND grade_id = (SELECT grade_id FROM grade WHERE grade_name = %s)
                AND subject_id = (SELECT subject_id FROM subject WHERE subject_abbr = %s)
            """, (grade_name, subject_abbr))

            db.commit()  
            self.deletestatusbar.showMessage("Scores deleted successfully", 3000)

        except Exception as e:
            db.rollback()  # Revert changes on error
            QMessageBox.critical(self,"AssessmentBoy",f"Database error: {e}")
    
    def deletegradescore(self):
        # Ask for confirmation
        confirm = QMessageBox.question(
            self,
            "AssessmentBoy",  # Title
            f"Are you sure you want to delete all\nscores for Grade {self.gradecombo.currentText()}?",  # Message
            QMessageBox.Yes | QMessageBox.No,  # Buttons
            QMessageBox.No  # Default button (avoids accidental deletion)
        )

        if confirm != QMessageBox.Yes:
            return  # Exit if user clicks "No"

        try:
            grade_name = self.gradecombo.currentText()

            # 1. Delete from score table
            cur.execute("""
                DELETE FROM score 
                WHERE grade_id = (SELECT grade_id FROM grade WHERE grade_name = %s)
                AND term_id = (SELECT term_id FROM term WHERE is_active = 1)
            """, (grade_name,))

            # 2. Delete from grand table
            cur.execute("""
                DELETE FROM grand 
                WHERE grade_id = (SELECT grade_id FROM grade WHERE grade_name = %s)
                AND term_id = (SELECT term_id FROM term WHERE is_active = 1)
            """, (grade_name,))

            # 3. Delete from total table
            cur.execute("""
                DELETE FROM total 
                WHERE grade_id = (SELECT grade_id FROM grade WHERE grade_name = %s)
                AND term_id = (SELECT term_id FROM term WHERE is_active = 1)
            """, (grade_name,))

            db.commit()
            QMessageBox.information(
                self,
                "AssessmentBoy",
                f"Scores for Grade {grade_name}\ndeleted successfully!",
                QMessageBox.Ok
            )

        except Exception as e:
            db.rollback()
            QMessageBox.critical(
                self,
                "AssessmentBoy - Error",
                f"Failed to delete scores:\n{str(e)}",
                QMessageBox.Ok
            )
    def deletelsubject(self):
        try:
            learner_id = int(self.delfield.text())
            
            # Get learner details
            cur.execute("SELECT first, second, surname FROM learner WHERE learner_id=%s", (learner_id,))
            name = cur.fetchone()
            if not name:
                QMessageBox.warning(self, "AssessmentBoy", "Learner not found")
                return
                
            full_name = f"{name[0]} {name[1]} {name[2]}"
            
            # Confirmation dialog
            resp = QMessageBox.question(
                self,
                "AssessmentBoy",
                f"Are you sure you want to delete\n{self.subjectcombo.currentText()} for {full_name}?",
                QMessageBox.Yes | QMessageBox.No
            )
            if resp != QMessageBox.Yes:
                return

            # SOFT DELETE: Update score to 0 instead of deleting
            cur.execute("""
                UPDATE score 
                SET subject_score = 0 
                WHERE learner_id = %s
                AND subject_id = (SELECT subject_id FROM subject WHERE subject_abbr = %s)
                AND term_id = (SELECT term_id FROM term WHERE is_active = 1)
            """, (learner_id, self.subjectcombo.currentText()))

            # Update grand total (sum will now exclude this subject)
            cur.execute("""
                SELECT COALESCE(SUM(subject_score), 0) 
                FROM score 
                WHERE learner_id = %s 
                AND term_id = (SELECT term_id FROM term WHERE is_active = 1)
            """, (learner_id,))
            new_grandtotal = cur.fetchone()[0]

            cur.execute("""
                UPDATE grand 
                SET grandtotal = %s
                WHERE learner_id = %s 
                AND term_id = (SELECT term_id FROM term WHERE is_active = 1)
            """, (new_grandtotal, learner_id))

            db.commit()
            QMessageBox.information(self, "AssessmentBoy", "Subject score delete successfully!")
            self.delfield.clear()
        except ValueError:
            QMessageBox.critical(self, "AssessmentBoy", "Invalid learner ID")
        except Exception as e:
            db.rollback()
            QMessageBox.critical(self, "AssessmentBoy", f"Error: {str(e)}")
            
#deleting all learning areas for a learner
    def deletelall(self):
        try:
            # Get learner ID from input field
            learner_id = int(self.delfield.text())
            
            # Access learner details
            cur.execute("""
                SELECT first, second, surname FROM learner
                WHERE learner_id = %s
            """, (learner_id,))
            name = cur.fetchone()
            
            if not name:
                QMessageBox.warning(self, "AssessmentBoy", "Learner not found")
                return
                
            full_name = f"{name[0]} {name[1]} {name[2]}"
            
            # Confirmation dialog
            resp = QMessageBox.question(
                self,
                "AssessmentBoy",
                f"Are you sure you want to CLEAR ALL\nSCORES for {full_name}?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No  # Default to 'No' for safety
            )
            
            if resp != QMessageBox.Yes:
                return

            # Get active term
            cur.execute("SELECT term_id FROM term WHERE is_active = 1")
            term_id = cur.fetchone()[0]

            # Soft delete all subject scores (set to 0)
            cur.execute("""
                UPDATE score 
                SET subject_score = 0 
                WHERE learner_id = %s
                AND term_id = %s
            """, (learner_id, term_id))

            # Update grand total to 0
            cur.execute("""
                UPDATE grand 
                SET grandtotal = 0
                WHERE learner_id = %s 
                AND term_id = %s
            """, (learner_id, term_id))

            db.commit()
            
            # Refresh UI if needed
            self.delfield.clear()
            QMessageBox.information(
                self, 
                "AssessmentBoy", 
                f"All scores for {full_name} cleared\nsuccessfully!",
                QMessageBox.Ok
            )
            
        except ValueError:
            QMessageBox.critical(self, "AssessmentBoy", "Please enter a valid learner ID")
        except Exception as e:
            db.rollback()
            QMessageBox.critical(
                self, 
                "AssessmentBoy", 
                f"Error clearing scores:\n{str(e)}"
            )
        
    def tohome(self):
        screen=mainwin()
        widget.addWidget(screen)
        widget.addWidget(screen)
        widget.setCurrentIndex(widget.currentIndex()+1)
    
class adminpanel(QMainWindow):
    def __init__(self):
        super().__init__()
        loadUi("admin.ui", self)
        self.homebutton.clicked.connect(self.tohome)
        # self.testcombo.currentTextChanged.connect(self.saveassessment)
        self.adminchange.clicked.connect(self.saveassessment)
    
    def saveassessment(self):
        test=self.testcombo.currentText()
        path="D:/TONNIEGIFTED/Documents/programs/AssessmentBoy/term.text"
        with open(path,"w")as file:
            file.write(test)
        self.adminstatusbar.showMessage("Changes Saved Successfully",3000)
      
        
    def tohome(self):
        screen=mainwin()
        widget.addWidget(screen)
        widget.addWidget(screen)
        widget.setCurrentIndex(widget.currentIndex()+1)


# Step 4: Launch app
window = QApplication(sys.argv)
screen = mainwin()
widget = QtWidgets.QStackedWidget()
widget.addWidget(screen)

# Make sure your setFixedSize works correctly with High DPI
widget.setFixedSize(480,375)
widget.setWindowTitle("AssessmentBoy")
widget.show()
sys.exit(window.exec_())
