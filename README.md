Commerce-Courses-Mass-Mailer
============================

Commerce Courses Mass Mailer is an application that can be used to send text emails to students registered for various Commerce courses. The application works by searching through objects on the Novell Network Tree of the University of Cape Town (UCT) network via LDAP (Light-weight Directory Access Protocol) calls, matching Course Groups on the user inputted filter, and extracting a list of student numbers from these located course groups. Once the student number list has been extracted, a list of email addresses is composed and the inputted text email is mailed to each address using UCT's mail server.  Created by Craig Lotter, October 2005
