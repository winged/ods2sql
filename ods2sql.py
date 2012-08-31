#!/usr/bin/python

import re
import sys
import copy
import zipfile
import string
import xml.parsers.expat

def get_inputfile():
	"""Returns a ZipFile object of the input.

	TODO: implement this to use ARGV, stdin etc as needed

	"""
	return zipfile.ZipFile('/dev/stdin', "r")

def get_outfile():
	"""Returns a filehandle to write SQL to.

	Could be stdout, or some other files, depnding on ARGV

	"""
	return sys.stdout


def get_xmldata():
	"""Returns a string containing the input file's content.xml (unparsed)
	"""
	ziparchive = get_inputfile()
	xmldata = ziparchive.read("content.xml")
	ziparchive.close()

	return xmldata

class Element(list):
	def __init__(self, name, attrs):
		self.name = name
		self.attrs = attrs

	@staticmethod
	def only_known(entry):
		if (entry.__class__ == Unknown):    return False
		if (entry.__class__ == StyleInfo):  return False
		if (entry.__class__ == Column):     return False

		return True

	def children(self):
		return list(filter(Element.only_known, self))

	def cleanup(self):
		while True:
			if len(self) <= 0: break
			if isinstance(self[-1], Element) and not self[-1].isempty(): break
			if isinstance(self[-1], str): break
			self.pop()


	def isempty(self):
		return False

	def __repr__(self):

		name     = self.__class__.__name__
		attrs    = self.attrs.__repr__()
		
		return "%s(%s: %s)" % (name, attrs, self.children())

class Unknown(Element):
	pass

class Database(Element):
	pass

class Table(Element):
	def __repr__(self):
		return "Table %s (%s) %s\n\n" % (self.attrs['table:name'], self.typemap, self.children())

	def cleanup(self):
		super().cleanup()

		types = []
		for row in self.children():
			cells = row.children()
			tlen  = len(types)
			for i in range(len(cells)):
				if i >= tlen:
					types.append(cells[i].gettype())
				else:
					oldtype = types[i]
					newtype = cells[i].gettype()
					types[i] = self.bettertype(oldtype, newtype)

		self.typemap = types
		self.name    = self.attrs['table:name']

	def bettertype(self, oldtype, newtype):
		# we know: int, float, string.
		# string can contain all types, float can contain int. int can only be int.
		#     same,   same: same
		#     string, *   : string
		#     float,  int : float
		#     

		if oldtype == 'string':  return 'string'
		if newtype == 'string':  return 'string'

		if newtype == 'float':   return 'float' # float vs string already handled

		if newtype == oldtype:   return oldtype # int, int
		
		return 'string'

class Row(Element):
	def __repr__(self):
		return "\n\tRow%s" % self.children()

	def isempty(self):
		return len(self) == 0

class Column(Element):
	def __repr__(self):
		return ""
	pass

class Wrapper(Element):
	def __repr__(self):
		children = self.children()
		if(len(children)) == 0: return ""
		if(len(children)) == 1: return "%s" % children[0].__repr__()
		return "%s" % repr(children)
	pass

class Cell(Element):
	def isempty(self):
		return self.content() == ''

	def gettype(self):
		content = self.content()
		if content.isdecimal(): return 'int'
		if content.isnumeric(): return 'float'
		return 'string'


	def content(self):
		return "".join([repr(x) for x in self.children()])

	def __repr__(self):
		return self.content()
		return "%s:%s" % (self.attrs, self.content())

class StyleInfo(Element):
	pass

class Paragraph(Element):
	def cleanup(self):
		pass
	def __repr__(self):
		return "".join(self.children())

class TreeBuilder:
	def __init__(self):
		self.root = Database("db", None)
		self.path = [self.root]
	def start_element(self, name, attrs):
		clones = 1
		if  (name == 'table:table'            ): element = Table    (name, attrs)
		elif(name == 'table:table-row'        ): element = Row      (name, attrs)
		elif(name == 'table:table-column'     ): element = Column   (name, attrs)
		elif(name == 'text:p'                 ): element = Paragraph(name, attrs)
		elif(name == 'office:spreadsheet'     ): element = Wrapper  (name, attrs)
		elif(name == 'office:body'            ): element = Wrapper  (name, attrs)
		elif(name == 'office:document-content'): element = Wrapper  (name, attrs)
		elif(name == 'table:named-range'      ): element = Unknown  (name, attrs)
		elif(re.match(r'^style:', name)       ): element = StyleInfo(name, attrs)
		elif(name == 'table:table-cell'       ):
			element = Cell (name, attrs)
			if 'table:number-columns-repeated' in element.attrs:
				clones = int(element.attrs['table:number-columns-repeated'])

		else:
			#print("Unknown element: %s %s\n" % (name, attrs))
			element = Unknown(name, attrs)

		for r in range(clones):
			self.path[-1].append(element)

		self.path.append(element)
	def end_element(self, name):
		assert name == self.path[-1].name
		if self.path[-1].__class__ == Wrapper:
			# remove wrapper
			for child in self.path[-1]:
				self.path[-2].append(child)
				if self.path[-1] in self.path[-2]:
					self.path[-2].remove(self.path[-1])
		else:
			self.path[-1].cleanup()

		self.path.pop()
	def char_data(self, data):
		self.path[-1].append(data)

def parse():
	xmlstring = get_xmldata()

	# create parser and parsehandler
	parser = xml.parsers.expat.ParserCreate()
	treebuilder = TreeBuilder()
	# assign the handler functions
	parser.StartElementHandler  = treebuilder.start_element
	parser.EndElementHandler    = treebuilder.end_element
	parser.CharacterDataHandler = treebuilder.char_data

	# parse the data
	parser.Parse(xmlstring, True)

	return treebuilder.root

def showtree(node, prefix=""):
	print (prefix, node.name)
	for e in node:
		if isinstance(e, Element):
			showtree(e, prefix + "  ")
		else:
			print(prefix + "  ", e)


def char_iter(end):
	chars = [c for c in string.ascii_uppercase]
	first = True
	for i in range(0,end):
		if first:
			yield 'A'
			first = False
		else:
			rep = ''
			while i > 0:
				mod  =     i%26
				i    = int(i/26)
				rep += chars[mod]
			yield rep

def render(db):
	typemap = {
		'string': 'TEXT',
		'int':    'INTEGER',
		'float':  'DOUBLE'
	}

	for table in db.children():
		colnames = []
		print('CREATE TABLE "%s"(' % table.name)
		fields= []
		fields.append('    "_id" INTEGER NOT NULL PRIMARY KEY')
		for name, coltype in zip(char_iter(len(table.typemap)), table.typemap):
			colnames.append(name)
			fields.append('    "%s" %s' % (name, typemap[coltype]))
		print(",\n".join(fields))
		print(");\n")

		print("BEGIN TRANSACTION;")
		for row in table.children():
			values = ['NULL']+["'%s'"% c.content().replace("'", "\\'") for c in row.children()]
			cols = ",".join(['"_id"']+['"%s"'% x for x in colnames[0:len(values)-1]])
			vals = ",".join(values)
			insert = '    INSERT INTO "%s" (%s) VALUES (%s);' % (table.name, cols, vals)
			print(insert)
		print("COMMIT;")
			


render(parse())
#print(showtree(parse()))
