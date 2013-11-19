# Emanuele Ruffaldi
# Scuola Superiore Sant'Anna PERCRO 2013
#
# TODO: when disembed verify presence of similar file to check
import zipfile
import argparse,os,sys
import xml.etree.ElementTree as ET
import urlparse,urllib
import subprocess,shutil

movieextension = ["mov","avi","wmv","mp4","m4v"]

"""

	ppt/_rels/presentation.xml.rels

	ppt/slides/_rels/slide#.xml.rels

Tags Relationship inside RelationShips, always in pairs
	Embedded: @Id @Type @Target
	Where @Target points to the folder ../media/XXX

	When external Target is a full url: file://... and we have @TargetMode=External

	@Type is: http://schemas.microsoft.com/office/2007/relationships/media
	@Type is: http://schemas.openxmlformats.org/officeDocument/2006/relationships/video

Tool:
	Actions:
	- extract: gets the embedded
	- disembed: removes as embedded and links externally
	- embed: embeds
	- makerelative: makes all references relative to current folder, or provided path

	Selector:
	- all
	- by slide
	- ?

Note: file url relative
	file://localhost/absolute/path/to/file 
	file:///absolute/path/to/file

"""

def urlfile2path(name):
	return urllib.unquote(urlparse.urlparse(name).path)

def path2urlfile(x):
	return "file:///" + urllib.quote(x)

class Media:
	def __init__(self,target,size=0,crc32=0):
		if target.startswith("file://"):
			self.extpath = urlfile2path(target) # path as url
		elif target.startswith("http://") or target.startswith("https://"):
			raise Exception("Unsupported http target")
		else:
			self.extpath = None # no specified for others
		# TODO http case (?)
		self.name = target
		self.target = target # full file in PPTX, if any
		self.crc32 = crc32 
		self.size = size
		self.uses = []
	@property
	def filename(self):
		if self.extpath is None:
			return os.path.split(self.target)[1]
		else:
			return os.path.split(self.extpath)[1]
	@property
	def isExternal(self):
		return self.extpath is not None
	@property
	def isInternal(self):
		return self.extpath is None
	@property
	def mode(self):
		return self.extpath is None and "External" or "Internal"
	def setExternal(self,path):
		self.target = path2urlfile(path)
		self.extpath = name
		self.updateuses()
	def setInternal(self,medianame):
		self.target = "ppt/media/"+medianame
		self.extpath = None
		self.updateuses()
	def updateuses(self):
		isext = self.extpath is not None
		for u in self.uses:
			if not isext:
				del u.node.attrib["TargetMode"]
			else:
				u.node.attrib["TargetMode"] = "External"
			u.node.attrib["Target"] = self.target

class MediaUse:
	def __init__(self,slide,media,node):
		self.slide = slide # object Slide
		self.media = media # object Media
		self.node = node # ET node

class Slide:
	def __init__(self,name,root):
		self.name = name # name of file in PPTX
		self.root = root # ET root of the loaded rels file
		self.uses = [] # list of medias in slide as MediaUse

def scanppt(filename,slide=None):
	slides = []
	medias = []
	zf = zipfile.ZipFile(filename,"r")
	medias = [Media(x.filename,crc32=x.CRC,size=x.file_size) for x in zf.infolist() if x.filename.startswith("ppt/media/")]
	medias = dict([(m.name,m) for m in medias])

	xslides = [x for x in zf.namelist() if x.startswith("ppt/slides/_rels/")]
	if slide is not None:
		if slide < 1 or slide > len(slides):
			print "slide not present",slide
			return ([],[])
		xslides = [xslides[slide-1]]
	
	slides = []
	for s in xslides:
		inf = zf.read(s)
		root = ET.fromstring(inf)
		slide = Slide(s,root)
		slides.append(slide)
		for c in root:
			if c.tag.endswith("Relationship"):
				if c.attrib["Type"].endswith("/video") or c.attrib["Type"].endswith("/media"):
					tm = "Internal"
					if "TargetMode" in c.attrib:
						tm = c.attrib["TargetMode"]
					tgt = c.attrib["Target"]
					media = None
					if tm == "Internal":
						if tgt.startswith("../media/"):
							media = medias[tgt[len("../media/"):]]
						else:
							raise Exception("Unknown internal media: " + s + " " + tgt)
					else:
						media = medias.get(tgt)
						if media is None:
							media = Media(tgt,"External")
							medias[tgt] = media
					mu = MediaUse(slide,media,c)
					media.uses.append(mu)
					slide.uses.append(mu)
					#print c.attrib["Id"],c.attrib["Type"],c.attrib["Target"],("TargetMode" in c.attrib and c.attrib["TargetMode"] or "")
					
	medias = dict([(m.name,m) for m in medias.values() if len(m.uses) > 0])
	zf.close()
	return slides,medias


def zipensuretmp(target=None):
	if target is None:
		target = "zipmanip_tmp"
	if os.path.isdir(target):
		shutil.rmtree(target)
	os.mkdir(target)
	return target

def zipextract(name,files,targetfiles,etarget):
	target = zipensuretmp(os.path.join(etarget,"zipmanip_tmp"))
	print "extracing to",target
	r = subprocess.check_output(["unzip","-o",name] + files + ["-d",target])
	print "moving in place"
	for i in range(0,len(files)):
		fext = os.path.join(etarget,targetfiles[i])
		fin = os.path.join(target,os.path.files[i])
		shutil.move(fin,fext)

def zipadd(name,extfiles,infiles):
	# based on ZipFile because we need just to add it by appending
	zf = zipfile.ZipFile(name,"a")
	for i in range(0,len(extfiles)):
		fext = extfiles[i]
		fin = infiles[i]
		zf.write(fext,fin,zipfile.ZIP_STORED)
	zf.close()

def zipdel(name,files):
	r = subprocess.check_output(["zip","-d"]+files)
	return 

def zipupdateslides(name,slides):
	# could be implemented by first removing the files using "zip" and then appending them again using Python
	target = zipensuretmp()
	for s in slides:
		sd = os.path.join(target,os.path.split(s.name)[0])
		if not os.path.isdir(sd):
			os.makedirs(sd)
		open(os.path.join(target,s.name),"wb").write(ET.tostring(s.root))
	
	a = os.chdir(target)
	subprocess.check_output(["zip","-u","-m",os.path.abspath(name),"."])
	os.chdir(a)

def findmediaentry(medias,ext):
	for i in range(1,1000000):
		me = "media%d.%s" % (i,ext)
		if me not in medias:
			return me
	return "impossible"
if __name__ == "__main__":
	import sys

	parser = argparse.ArgumentParser(description='PPTX Video Manager')
	parser.add_argument('--slide',type=int,help="slide to process otherwise all")
	parser.add_argument('--extract',action="store_true",help="extracts the medias that are")
	parser.add_argument('--collect',action="store_true",help="extract and collect all medias, even external")
	parser.add_argument('--embed',action="store_true",help="embeds all the externals (whatever the extension)")
	parser.add_argument('--disembed',help="disembed all the externals to PPT file path",action="store_true")
	parser.add_argument('--fix',help="fixes the paths to the PPT directory if they are absolute",action="store_true")
	parser.add_argument('--list',help="list types and content",action="store_true")
	parser.add_argument('--rename',nargs=2,help="renames the local filename referenced to another name")
	parser.add_argument('--verify',help="Verify external",action="store_true")
	#parser.add_argument('--output',help="output path name of the PPT, otherwise modified. Only for single")
	parser.add_argument('input',nargs="+")

	args = parser.parse_args()


	#if len(args.input) > 1 and args.output != None:
	#	print "output only for specific file"
	#	sys.exit(0)


	for x in args.input:
		xp,xn = os.path.split(x)
		bp = os.path.abspath(xp)
		xnoe = os.path.splitext(x)[0]
		slides,medias = scanppt(x,slide=args.slide)

		print "**Slides"
		for s in slides:
			if len(s.uses) == 0:
				continue
			print "slide",s.name.split("/")[-1].replace(".xml.rels","")
			for u in s.uses:
				if u.media.target != "NULL":
					print "\tuses",u.media.target
		if len(medias) > 0:
			print "**Medias"
			for m in medias.values():
				if m.target != "NULL":
					print "media (used %d) %s" % (len(m.uses),m.target)

		if args.extract or args.collect:
			toextract = []
			toextractoutname = []
			tocopy = []
			for m in medias.values():
				if m.isInternal:
					toextractoutname.append(xnoe + "_" + m.filename)
					toextract.append(m.target)
					print "plan to extract",toextract[-1],"as",toextractoutname[-1]
				elif args.collect:
					ep,en = os.path.split(m.extpath)
					eap = os.path.abspath(ep)					
					# only if not already in same folder
					if eap != bp:
						tocopy.append(m.extpath)
						print "plan to copy",tocopy[-1],"to target"
					else:
						print "skip local file",m.extpath
			print "extracting files"
			zipextract(x,toextract,toextractoutname,bp)

			if len(tocopy) > 0:
				print "copying files to target directory"
				for x in tocopy:
					shutil.copyfile(x,os.path.join(bp,os.path.split(x)[1]))
		elif args.verify:
			for m in medias.values():
				if m.isExternal:
					if os.path.isfile(m.extpath) == 0:
						print "missing",m.extpath
					else:
						print "OK",m.extpath
		elif args.embed:
			toaddin = []
			toaddext = []
			for m in medias.values():
				if m.isExternal:
					ext = os.path.splitext(m.extpath)[1]
					extpath = m.extpath
					me = findmediaentry(medias,ext)
					print "plan to add",extpath,"as",me
					if not os.path.isfile(extpath):
						extpath = extpath.split("/")[-1].split("\\")[-1]
						print "\tfile missing remapped input as",extpath
					if not os.path.isfile(extpath):
						print "ERROR: file missing anyway",extpath
						continue
					toaddext.append(extpath)
					toaddin.append(me)
					m.setInternal(me)
			# TODO: optimize by one single add and update
			zipadd(x,toaddext,toaddin)
			zipupdateslides(x,slides)
		elif args.disembed:
			# pull them out
			toextract = []
			toextractoutname = []
			for m in medias.values():
				if m.isInternal:
					toextract.append(m.target)
					toextractoutname.append(xnoe+"_"+m.filename)
					print "plant to disembed",toextract[-1],"as",toextractoutname[-1]
					m.setExternal(os.path.join(bp,toextractoutname[-1]))
			zipextract(x,toextract,toextractoutname,bp)
			zipupdateslides(x,slides)
		elif args.fix:
			for m in medias.values():
				if m.isExternal:
					pp,b = os.path.split(m.extpath)
					pp = os.path.abspath(pp)
					if pp != bp:
						newextpath = os.path.join(bp,b)
						print "plan to relocate",m.extpath,"to",newextpath
						m.setExternal(newextpath)
						if os.path.isfile(newextpath):
							print "Warning: the new target media file is missing",newextpath
					else:
						print "skip",m.extpath
			zipupdateslides(x,slides)
		elif args.rename is not None and len(args.rename) == 2:
			ap0 = os.path.abspath(args.rename[0])
			ap1 = os.path.abspath(args.rename[1])
			oslides = []
			for m in medias.values():
				if m.isExternal:
					if m.extpath == ap0:
						print "renaming real file",args.rename[0],args.rename[1]
						os.rename(ap0,ap1)
						m.setExternal(ap1)
						oslides = [u.slide for u in m.uses]
						break
			if len(oslides) > 0:
				zipupdateslides(x,oslides)
			else:
				print "notfound",args.rename[0],"as external"
		elif args.list:
			for m in medias.values():
				print m.target,m.extpath,len(m.uses)

