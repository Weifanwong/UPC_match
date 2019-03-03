class MyClass(object):
	def my_func1(self):
		print('my_func1_in')

	def my_func2(self):
		self.my_func1()


def my_func1():
	print('my_func1_out')

myclass1 = MyClass();
myclass1.my_func2();