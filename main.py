"""
In gmail it's necessary to allow less secure applications to login:
https://support.google.com/mail/answer/7126229
"""

__author__ = 'Surrerstry'
__version__ = 1.0
__website__ = 'surrerstry.pl'

import smtplib
import logging
import tkinter
import tkinter.scrolledtext

from os.path import exists
from threading import Thread
from tkinter import messagebox
from tkinter.filedialog import askopenfilename
from importlib import reload
from inspect import currentframe, getframeinfo
from time import sleep
from collections import namedtuple
from random import uniform, shuffle

gui_exists = True
status_text = """\nCurrent status: {}\n\nLoaded targets: {}\n\nAlready sent: {}\nLeft to send: {}"""
current_status = 'Idle'
already_sent = 0

##### CONFIGURATION MENU #####

logging_level = 'Debug' # Info, Warning, Error, Critical
logging_level_connections = {
	'Debug':['Debug', 'Info', 'Warning', 'Error', 'Critical'],
	'Info':['Info', 'Warning', 'Error', 'Critical'],
	'Warning':['Warning', 'Error', 'Critical'],
	'Error':['Error', 'Critical'],
	'Critical':['Critical'],
	}
order_of_sending = 'One by one' # Mixed up
Delay = namedtuple('Delay', ['min', 'max'])
delay_in_second = Delay(min=0.1, max=0.5)

##### ------------------ #####


class log(object):

	def __generic_for_each_level__(msg, calling_place=None):

		if calling_place is not None:
			calling_place = getframeinfo(calling_place)
			data_to_save = '{} : {}:{}'.format(msg, calling_place.filename, calling_place.lineno)
		else:
			data_to_save = '{}'.format(msg)
		
		return data_to_save


	@staticmethod
	def debug(msg, calling_place=None):
		if 'Debug' not in logging_level_connections[logging_level]:
			return
		res = log.__generic_for_each_level__(msg, calling_place)
		logger.debug(res)

	@staticmethod
	def info(msg, calling_place=None):
		if 'Info' not in logging_level_connections[logging_level]:
			return
		res = log.__generic_for_each_level__(msg, calling_place)
		logger.info(res)

	@staticmethod
	def warning(msg, calling_place=None):
		if 'Warning' not in logging_level_connections[logging_level]:
			return
		res = log.__generic_for_each_level__(msg, calling_place)
		logger.warning(res)

	@staticmethod
	def error(msg, calling_place=None):
		if 'Error' not in logging_level_connections[logging_level]:
			return
		res = log.__generic_for_each_level__(msg, calling_place)
		logger.error(res)

	@staticmethod
	def critical(msg, calling_place=None):
		if 'Critical' not in logging_level_connections[logging_level]:
			return
		res = log.__generic_for_each_level__(msg, calling_place)
		logger.critical(res)


class GUI(object):

	def __init__(self, master):
		master.title('Massmailing System')

		self.master = master

		self.min_width_of_window = 860
		self.min_height_of_window = 650

		self.let_send = False
		self.senders_to_import = []

		self.left_to_send = 0
		self.loaded_emails = 0

		self.tls_val = 0

		global logging_level, order_of_sending, delay_in_second
		self.logging_level = logging_level
		self.order_of_sending = order_of_sending
		self.delay_in_second = delay_in_second

		self.check_var = tkinter.IntVar()

	def menu_bar(self):
		self.menu = tkinter.Menu(self.master)
		self.master.config(menu=self.menu)

	def sub_menu_1(self):
		self.sub_menu_1 = tkinter.Menu(self.menu)
		self.menu.add_cascade(label='Program', menu=self.sub_menu_1)
		
		self.sub_menu_1.add_command(label="Configuration", command=lambda: self.__class__.configuration(self))
		self.sub_menu_1.add_command(label="Exit", command=self.master.quit)

	def sub_menu_2(self):
		self.sub_menu_2 = tkinter.Menu(self.menu)
		self.menu.add_cascade(label='Import', menu=self.sub_menu_2)
		
		self.sub_menu_2.add_command(label="Target Emails (from file)", command=lambda: self.__class__.import_targets(self))
		self.sub_menu_2.add_command(label="Sender Emails (from file)", command=lambda: self.__class__.import_senders(self))

	def status_bar(self):
		status = tkinter.Label(self.master, text='Massmailing 1.0 | Surrerstry.pl | Ⓒ All Rights Reserved', bd=1, relief=tkinter.SUNKEN)
		status.pack(side=tkinter.BOTTOM, fill=tkinter.X)

		self.email_content = tkinter.scrolledtext.ScrolledText(self.master, height=20, width=94, bd=4, padx=5, pady=5)
		self.email_content.insert(tkinter.END, """Hello,

Here you should paste content of your email.

Regards.
												""")
		
		self.email_content.pack(side=tkinter.BOTTOM, anchor=tkinter.W)

		label_between = tkinter.Label(self.master, height=2, text='{}Content of the emails:'.format(' '*67))
		label_between.pack(side=tkinter.BOTTOM, anchor=tkinter.W)

	def set_geometry(self):

		self.width_of_screen = root.winfo_screenwidth()
		self.height_of_screen = root.winfo_screenheight()

		self.master.geometry('{}x{}+{}+{}'.format(self.min_width_of_window, self.min_height_of_window, (self.width_of_screen // 2) - (self.min_width_of_window // 2), (self.height_of_screen // 2) - (self.min_height_of_window // 2) ))
		self.master.minsize(self.min_width_of_window, self.min_height_of_window)
		self.master.maxsize(self.min_width_of_window, self.min_height_of_window)

		self.main_frame = tkinter.Frame(self.master, width=self.min_width_of_window, heigh=self.min_height_of_window)
		self.main_frame.pack()

	def email_fields(self):
	
		label = tkinter.Label(self.master, heigh=1, bd=3, text='Recipients Emails:{}Senders Emails:{}'.format(' '*60, ' '*48))
		label.pack(side=tkinter.TOP)

		self.target_field = tkinter.scrolledtext.ScrolledText(self.master, width=45, height=20, bd=3, padx=5, pady=5)
		self.target_field.insert(tkinter.END, "recipient_1@sth.com"+
											  "\nrecipient_2@sth.com"+
											  "\nrecipient_3@sth.com"+
											  "\nrecipient_4@sth.com")
		self.target_field.pack(side=tkinter.LEFT)

		self.sender_field = tkinter.scrolledtext.ScrolledText(self.master, width=45, height=20, bd=3, padx=5, pady=5)
		self.sender_field.insert(tkinter.END, "sender_1@sth.com:password:smtp.domain.com:587"+
											  "\nsender_2@sth.com:password:smtp.domain.com:587")
		self.sender_field.pack(side=tkinter.LEFT)

	def start_and_stop_buttons(self):

		space_label = tkinter.Label(self.master, text='', height=1)
		space_label.pack(side=tkinter.TOP)

		start_button = tkinter.Button(self.master, text='Start mailing', bd=1, width=9, command=lambda: self.__class__.start_sending_emails(self))
		start_button.pack(side=tkinter.TOP)

		space_label = tkinter.Label(self.master, text='', height=1)
		space_label.pack(side=tkinter.TOP)

		stop_button = tkinter.Button(self.master, text='Stop mailing', bd=1, width=9, command=lambda: self.__class__.stop_sending_emails(self))
		stop_button.pack(side=tkinter.TOP)

		self.statistic_label = tkinter.Label(self.master, text=status_text.format('Idle', self.loaded_emails, already_sent, self.left_to_send), height=10)
		self.statistic_label.pack(side=tkinter.TOP)


	def configuration(self):
		# ----- Position of configuration menu -----
		self.configuration = tkinter.Toplevel(width=200)
		self.configuration.title("Configuration")

		self.configuration.lift()

		self.configuration.geometry('{}x{}+{}+{}'.format(self.min_width_of_window // 2, self.min_height_of_window // 2, ((self.width_of_screen // 2) - (self.min_width_of_window // 2)) * 2, ((self.height_of_screen // 2) - (self.min_height_of_window // 2)) * 2 ))

		self.configuration.minsize(self.min_width_of_window // 2, self.min_height_of_window // 2)
		self.configuration.maxsize(self.min_width_of_window // 2, self.min_height_of_window // 2)

		# ---------- OPTIONS IN CONFIGURATION ----------

		# ----- level of logging ----- 
		tkvar = tkinter.StringVar(self.configuration)

		choices = ['Debug','Info','Warning','Error','Critical']
		tkvar.set(self.logging_level)

		self.popupMenu = tkinter.OptionMenu(self.configuration, tkvar, *choices, command=self.change_debugging_level)
		tkinter.Label(self.configuration, text="\nLevel of logging").pack(side=tkinter.TOP)
		self.popupMenu.pack(side=tkinter.TOP)

		# ----- Order of sending -----
		tkvar = tkinter.StringVar(self.configuration)

		choices = ['One by one','Mixed up']
		tkvar.set(self.order_of_sending)

		self.popupMenu2 = tkinter.OptionMenu(self.configuration, tkvar, *choices, command=self.change_order_of_sending)
		tkinter.Label(self.configuration, text="\nOrder of sending").pack(side=tkinter.TOP)
		self.popupMenu2.pack(side=tkinter.TOP)

		# ----- Delay in seconds -----
		msg = tkinter.Message(self.configuration, text='\nDelay in seconds (min, max)', width=200)
		msg.pack(side=tkinter.TOP)

		self.sleeps_frame = tkinter.Frame(self.configuration)
		self.entry_1_delay = tkinter.Entry(self.sleeps_frame, width=5)
		self.entry_2_delay = tkinter.Entry(self.sleeps_frame, width=5)

		self.entry_1_delay.insert(tkinter.END, delay_in_second.min)
		self.entry_2_delay.insert(tkinter.END, delay_in_second.max)

		self.entry_1_delay.pack(side=tkinter.LEFT)
		self.entry_2_delay.pack(side=tkinter.LEFT)

		self.sleeps_frame.pack()

		# ----- TLS usage -----
		self.tls_frame = tkinter.Frame(self.configuration)

		msg = tkinter.Message(self.tls_frame, text='')
		msg.pack(side=tkinter.TOP)

		self.tls = tkinter.Checkbutton(self.tls_frame, text='Use TLS', variable=self.check_var)
		self.tls.pack(side=tkinter.TOP)

		msg = tkinter.Message(self.tls_frame, text='')
		msg.pack(side=tkinter.TOP)

		self.tls_frame.pack(side=tkinter.TOP)

		# ----- Apply button -----
		apply_button = tkinter.Button(self.tls_frame, text='Apply', bd=1, width=9, command=lambda: self.__class__.save_configuration(self))
		apply_button.pack(side=tkinter.TOP)

	def change_debugging_level(self, val):
		self.logging_level = val

	def change_order_of_sending(self, val):
		self.order_of_sending = val

	def save_configuration(self):

		try:
			self.msg_apply.destroy()
		except:
			pass

		global logging_level, order_of_sending, delay_in_second
		logging_level = self.logging_level
		order_of_sending = self.order_of_sending

		delay_in_second = Delay(min=self.entry_1_delay.get(), max=self.entry_2_delay.get())

		if self.tls_val != self.check_var.get():
			self.tls_val = self.check_var.get()

		self.check_var.set(self.tls_val)

		self.msg_apply = tkinter.Message(self.tls_frame, text='Configuration saved', width=200, fg='green')
		self.msg_apply.pack(side=tkinter.TOP)
		
	def import_senders(self):

		if current_status != 'Idle':
			messagebox.showinfo("Import Error", "Cannot import while sending in progress.")
			return 

		self.filename_senders_to_import = askopenfilename()

		if not isinstance(self.filename_senders_to_import, str) or not exists(self.filename_senders_to_import):
			return "User not choosed file, or file not available anymore..."

		with open(self.filename_senders_to_import) as file:
			self.senders_to_import = file.readlines()

		self.sender_field.delete(1.0, tkinter.END)
		self.sender_field.insert(tkinter.END, ''.join(self.senders_to_import))

		self.tmp_list_of_senders = self.sender_field.get(1.0, tkinter.END).split()
		self.tmp_list_of_senders_idx = 0
		self.tmp_list_of_senders_amount = len(self.tmp_list_of_senders)

	def import_targets(self):

		if current_status != 'Idle':
			messagebox.showinfo("Import Error", "Cannot import while sending in progress.")
			return

		self.filename_targets_to_import = askopenfilename()

		if not isinstance(self.filename_targets_to_import, str) or not exists(self.filename_targets_to_import):
			return "User not choosed file, or file not available anymore..."

		with open(self.filename_targets_to_import) as file:
			self.targets_to_import = file.readlines()

		self.loaded_emails = len(self.targets_to_import)
		self.left_to_send = self.loaded_emails

		self.statistic_label.config(text=status_text.format('Idle', self.loaded_emails, already_sent, self.left_to_send))

		self.target_field.delete(1.0, tkinter.END)
		self.target_field.insert(tkinter.END, ''.join(self.targets_to_import))
 
	def start_sending_emails(self):
		self.let_send = True
		self.statistic_label.config(text=status_text.format('Sending...', self.loaded_emails, already_sent, self.left_to_send))
		global current_status
		current_status = 'Sending...'

		self.loaded_emails = len(self.target_field.get(1.0, tkinter.END).split())
		self.left_to_send = self.loaded_emails

		self.statistic_label.config(text=status_text.format('Idle', self.loaded_emails, already_sent, self.left_to_send))

	def stop_sending_emails(self):
		self.let_send = False
		self.statistic_label.config(text=status_text.format('Idle', self.loaded_emails, already_sent, self.left_to_send))
		global current_status
		current_status = 'Idle'

	def sending_loop(self):
		self.tmp_list_of_senders = self.sender_field.get(1.0, tkinter.END).split()
		self.tmp_list_of_senders_idx = 0
		self.tmp_list_of_senders_amount = len(self.tmp_list_of_senders)
		
		global order_of_sending, delay_in_second

		while True:
			while self.let_send:
				# load mails from scrollbox
				tmp_list_of_targets = self.target_field.get(1.0, tkinter.END).split()

				if order_of_sending != 'One by one':
					shuffle(tmp_list_of_targets)

				if self.tmp_list_of_senders_idx >= self.tmp_list_of_senders_amount:
					self.tmp_list_of_senders_idx = 0

				send_from = self.tmp_list_of_senders[self.tmp_list_of_senders_idx].split(':')

				if len(tmp_list_of_targets) == 0:
					# list empty -> go to waiting loop
					break

				tmp_to_send = tmp_list_of_targets.pop(0)

				# clear scrollbox
				self.target_field.delete(1.0, tkinter.END)
				# put to scrollbox one element less
				self.target_field.insert(tkinter.END, '\n'.join(tmp_list_of_targets))

				self.tmp_list_of_senders_idx += 1

				global already_sent
				already_sent += 1
				self.left_to_send -= 1

				self.statistic_label.config(text=status_text.format('Sending...', self.loaded_emails, already_sent, self.left_to_send))

				content_of_the_email = self.email_content.get(1.0, tkinter.END)

				print(send_from[2], int(send_from[3]), send_from[0], send_from[1], self.tls_val)
				print(send_from[0], tmp_to_send)
				server_object = login_to_server(send_from[2], int(send_from[3]), send_from[0], send_from[1], tls=self.tls_val)
				send_email(server_object, send_from[0], tmp_to_send, content_of_the_email)
				quit_connection(server_object)
				
				sleep(uniform(float(delay_in_second.min), float(delay_in_second.max)))

				if gui_exists == False:
					return

			sleep(1) # wait on let.send

			if gui_exists == False:
				return 


def set_logging(format, filename='logs.log', level=logging.DEBUG):
	logging.basicConfig(filename=filename,
						level=level,
						format=format)

def login_to_server(smtp_server, smtp_port, user_email, password, tls=False):
	
	log.debug('Login to server: {}:{}:{}:{}:TLS={}'.format(smtp_server, smtp_port, user_email, password, tls), currentframe())

	server = smtplib.SMTP(smtp_server, smtp_port)
	server.connect(smtp_server, smtp_port)

	if tls:
		server.starttls()

	server.login(user_email, password)

	return server

def send_email(server_object, user_email, target_email, content):

	log.info('Send email: from:{}, to:{}'.format(user_email, target_email), currentframe())
	log.info('Content:{}'.format(content), currentframe())

	server_object.sendmail(user_email, target_email, content)

def quit_connection(server_object):

	log.debug('Quit connection', currentframe())

	server_object.quit()

def create_gui(root):
	gui = GUI(root)
	
	gui.menu_bar()
	gui.sub_menu_1()
	gui.sub_menu_2()
	
	gui.status_bar()
	
	gui.email_fields()
	gui.start_and_stop_buttons()
	
	gui.set_geometry()
	sending_loop_thread = Thread(target=gui.sending_loop)
	sending_loop_thread.start()

	return sending_loop_thread


if __name__ == '__main__':

	set_logging('%(message)s')

	# log separation line
	logger = logging.getLogger()
	logger.info('{}'.format('-'*80))

	reload(logging)

	# set detailed format
	set_logging('%(levelname)s:%(asctime)s - %(message)s')
	
	logger = logging.getLogger()
	log.info('Application started', currentframe())

	root = tkinter.Tk()
	sending_loop_thread = create_gui(root)

	root.mainloop()
	gui_exists = False
	sending_loop_thread.join()

	log.info('Application closed', currentframe())
	exit(0)
