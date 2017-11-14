Attribute VB_Name = "test_logger"
Option Explicit

Public Function testMailLogger()

Dim obj_logger As LoggerMail
Dim Message As MSG


Set obj_logger = New LoggerMail

obj_logger.init "test", log4VBA.ERRO, "test_destination"
obj_logger.mailAddress = "rostislav.jirasek@lego.com"

log4VBA.init
log4VBA.add_logger obj_logger



Set Message = New MSG
log4VBA.error "test_destination", Message.source("testing_module", "testing-function").text("Toto je testovací zpráva, která nemá žádný význam.")
Set Message = Nothing


End Function

Public Function testFileLogger()

Dim obj_logger As LoggerFile
Dim Message As MSG


Set obj_logger = New LoggerFile

obj_logger.init "test", log4VBA.ERRO, "test_destination"
obj_logger.logFilePath = "C:\Users\czJiRost\Desktop\log-format-example.xlsx"



log4VBA.init
log4VBA.add_logger obj_logger



Set Message = New MSG
log4VBA.error "test_destination", Message.source("testing_module", "testing-function").text("Problém je mezi židlí a klávesnicí")
Set Message = Nothing


End Function

