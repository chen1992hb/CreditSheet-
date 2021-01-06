#coding = utf-8
from appium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support.wait import WebDriverWait
import time
import ngender


str =ngender.guess('刘慧')

print(str)
