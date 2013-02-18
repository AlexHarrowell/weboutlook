"""
Microsoft Outlook Web Access scraper

Retrieves full, raw e-mails from Microsoft Outlook Web Access by
screen scraping. Can do the following:

	* Log into a Microsoft Outlook Web Access account with a given username
      and password.
	* Retrieve all e-mail IDs from any given folder
	* Retrieve the full, raw source of the e-mail with a given ID
	* Delete an e-mail with a given ID (technically, move it to the "Deleted
      Items" folder).

The main class you use is OutlookWebScraper. See the docstrings in the code
and the "sample usage" section below.

This module does do caching! It caches your session so that you only have to log
in once. It also caches the message IDs for each folder, and the content of any message you retrieve, to save on HTTP requests. the get_folder() and inbox() methods let you specify refresh=True to flush the cache, and you can call flush_cache() to remove all cached information.

Updated by Greg Albrecht <gba@gregalbrecht.com>
Ported to use mechanize, handle form-based authentication, handle folders with more than one page, and cache requests by Alexander Harrowell <a.harrowell@gmail.com>
Based on http://code.google.com/p/weboutlook/ by Adrian Holovaty <holovaty@gmail.com>.
"""

# Documentation / sample usage:
#
# # Throws InvalidLogin exception for invalid username/password.
# >>> s = OutlookWebScraper('https://webmaildomain.com', 'username', 'invalid password')
# >>> s.login()
# Traceback (most recent call last):
#     ...
# scraper.InvalidLogin
#
# >>> s = OutlookWebScraper('https://webmaildomain.com', 'username', 'correct password')
# >>> s.login()
#
# # Display IDs of messages in the inbox.
# >>> s.inbox()
# ['/Inbox/Hey%20there.EML', '/Inbox/test-3.EML']
#
# # Display IDs of messages in the "sent items" folder.
# >>> s.get_folder('sent items')
# ['/Sent%20Items/test-2.EML']
#
# # Display the raw source of a particular message.
# >>> print s.get_message('/Inbox/Hey%20there.EML')
# [...]
#
# # Delete a message.
# >>> s.delete_message('/Inbox/Hey%20there.EML')

# Copyright (C) 2006 Adrian Holovaty <holovaty@gmail.com>
#
# This program is free software; you can redistribute it and/or modify it under
# the terms of the GNU General Public License as published by the Free Software
# Foundation; either version 2 of the License, or (at your option) any later
# version.
#
# This program is distributed in the hope that it will be useful, but WITHOUT
# ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS
# FOR A PARTICULAR PURPOSE. See the GNU General Public License for more
# details.
#
# You should have received a copy of the GNU General Public License along with
# this program; if not, write to the Free Software Foundation, Inc., 59 Temple
# Place, Suite 330, Boston, MA 02111-1307 USA

import socket, urllib, urlparse
import logging
from logging.handlers import *
from mechanize import Browser

__version__ = '0.1.4'
__author__ = 'Greg Albrecht <gba@gregalbrecht.com>, Alexander Harrowell <a.harrowell@gmail.com>'

logger = logging.getLogger('weboutlook')
logger.setLevel(logging.INFO)
consolelogger = logging.StreamHandler()
consolelogger.setLevel(logging.INFO)
logger.addHandler(consolelogger)

socket.setdefaulttimeout(15)

class InvalidLogin(Exception):
	pass

class RetrievalError(Exception):
	pass

class OutlookWebScraper():
	def __init__(self, domain, username, password):
		logger.debug(locals())
		self.domain = domain
		self.username, self.password = username, password
		self.is_logged_in = False
		self.base_href = None
		self.browser = Browser() # home to the mechanize headless browser-like entity
		self.folder_cache = {} # a dict for storing folder names and lists of messages
		self.message_cache = {} # a dict for storing message IDs and content

	def add_to_cache(self, folder_name=None, message_urls=None, msgid=None, payload=None):
		if folder_name: # simple - stores the message IDs from a folder with the folder name as a key
			self.folder_cache[folder_name] = message_urls
			return self.folder_cache[folder_name]
			
		if msgid: # even simpler - store the raw e-mail under the message id
			self.message_cache[msgid] = payload
			return self.message_cache[msgid]

	def find_in_cache(self, folder_name=None, msgid=None):
		if msgid: # quick key-in-d lookup 
			if msgid in self.message_cache:
				return self.message_cache[msgid]
		if folder_name:
			if folder_name in self.folder_cache:
				return self.folder_cache[folder_name]
		else: # i.e. no message of that id or no folder of that name
			return False

	def remove_from_cache(self, folder_name=None, msgid=None):
		if msgid:
			cache = find_in_cache(msgid=msgid) # search for the message
			if cache:
				del self.message_cache[msgid] # remove it
				if not folder_name: # if that's it..				
					return True
			# if a folder name was provided, or no message found, though..
		if folder_name:
			cache = find_in_cache(folder_name=folder_name) # search for a folder
			if cache:
				del self.folder_cache[folder_name] # if found, kill it
				for id in cache:
					del self.message_cache[id] # and remove the messages
				return True
			else:
				return False
		else: # nothing was found
			return False
			
	def flush_cache(self):
		self.folder_cache = {} # really simple
		self.message_cache = {}
		return True
			
	def login(self):
		logger.debug(locals())
		destination = urlparse.urljoin(self.domain, 'exchange/')
		self.browser.add_password(destination, self.username, self.password)
		self.browser.open(destination) # it should just work with basic auth, but let's deal with form as well
		self.browser.select_form('logonForm')
		self.browser['username'] = self.username
		self.browser['password'] = self.password
		self.browser.submit()
		if 'You could not be logged on to Outlook Web Access' in self.browser.response().read():
			raise InvalidLogin
		m = self.browser.links().next()
		if not m.base_url:
			raise RetrievalError, "Couldn't find <base href> on page after logging in."
		self.base_href = m.base_url
		self.is_logged_in = True
	
	def inbox(self, refresh=None):
		"""
		Returns the message IDs for all messages in the
		Inbox, regardless of whether they've already been read. setting refresh forces an update.
		"""
		logger.debug(locals())
		if refresh: # refresh kwarg. if set, forces a fresh load of the folder
			return self.get_folder('Inbox', refresh=True)
		else:
			return self.get_folder('Inbox')

	def get_folder(self, folder_name, refresh=None):
		"""
		Returns the message IDs for all messages in the
		folder with the given name, regardless of whether the messages have
		already been read. The folder name is case insensitive. setting refresh forces an update. if not set, cached messages
		will be returned if they exist.
		"""
		logger.debug(locals())
		if not refresh: # look in the cache
			message_urls = find_in_cache(folder_name=folder_name)
			if message_urls:
				return message_urls
		# if not found or not used, proceed
		if not self.is_logged_in: 
			self.login()
		url = self.base_href + urllib.quote(folder_name) + '/?Cmd=contents'
		self.browser.open(url)
		message_urls = [link.url for link in self.browser.links() if '.EML' in link.url]
		if '&nbsp;of&nbsp;1' in self.browser.response().read(): # test for multiple pages
			add_to_cache(folder_name=folder_name, message_urls=message_urls) # cache it			
			return message_urls
		else:
			last_message_urls = message_urls # if you ask for an out-of-range page, you get the last page
			flag = True
			while flag: # page through the folder, testing to see if we've got the last one yet
				next_url = [link.url for link in self.browser.links() if 'Next Page' in link.text][0]	
				self.browser.open(next_url) #find and use next page url
				murls = [(link.url).replace((self.base_href), '') for link in self.browser.links() if '.EML' in link.url] #extract data
				if murls == last_message_urls:
					flag = False # check to see if identical with the last one
				last_message_urls = murls
				message_urls.extend(murls) # build the list
			add_to_cache(folder_name=folder_name, message_urls=message_urls)	
			return message_urls # if you get the same page you've finished
	    
	def get_message(self, msgid):
		"Returns the raw e-mail for the given message ID."
		logger.debug(locals())
		payload = find_in_cache(msgid=msgid) #check the cache
		if payload:
			return payload
		if not self.is_logged_in:
			self.login()
		# Sending the "Translate=f" HTTP header tells Outlook to include
		# full e-mail headers. Figuring that out took way too long.
		self.browser.addheaders = [('Translate', 'f')]
		payload = self.browser.open(self.base_href + msgid + '?Cmd=body')
		add_to_cache(msgid=msgid, payload=payload) #cache the results
		return payload
		
	def delete_message(self, msgid):
		"Deletes the e-mail with the given message ID."
		logger.debug(locals())
		if not self.is_logged_in: 
			self.login()
		msgid = msgid.replace('?Cmd=open', '')
		delete = self.browser.open(self.base_href + msgid, urllib.urlencode({
		    'MsgId': msgid,
		    'Cmd': 'delete',
		    'ReadForm': '1',   
		 }))
		remove_from_cache(msgid=msgid) #and remove from cache
		return delete

