set(['Erika Tsutsumi', 'Heidi Nafis', 'Jessica Noglows', 'Robert McMullen', 'Jordyn Burger', 'Myles Cooper', 'Eric VanWyk', 'Emma', 'Erin McCusker', 'Scott McClure', 'Chris Lee', 'Jessie Oehrlein', 'Victoria Preston', 'Shane M Skikne', 'Tiana Veldwisch', 'Sutee Dee', 'Stephen Longfield', 'Juliana Nazare', 'Victoriea Bird', 'Rose Higgins', 'Kevin Bretney', 'Katie (Rivard) Mazaitis', 'Adam Kenvarg', 'James Nee', 'Kimble McCraw', 'Janaki', 'Eric Schneider', 'Matthew Hill', 'Erika Swartz', 'Tom Chen', 'Andrew Barry', 'AVERY LOUIE', 'Kevin Simon', 'Kori Ryter', 'herbert', 'Graham Hooton', 'Nikki Lee', 'Sean McBride', 'Rachel', 'Adriana Garties', 'Amy Whitcombe', 'Casey Monahan', 'Nicole Rifkin', 'Matt Ritter', 'Chris Mark', 'Kelcy Adamec', 'Neal Singer', 'Rachel Bobbins', 'Alex Wheeler', 'Alyshia Olsen', 'Alex Dorsk', 'Sarah Seko', 'Caroline Condon', 'Jen Wei', 'Erika Weiler', 'Joe Kendall', 'Aman Kapur', 'Andy Pethan', 'Christina Nguyen', 'Tara Krishnan', 'Jessica Anderson'])

comma = '%%%'

import random

def makeStudents():
  f = open('mentorselection_all2.tsv')
  students = dict()
  students = makeStruct(f, students)
  return students

# def makeAlums(students):
#   f = open('mentorselection_recentalumns.tsv')
#   students = makeStruct(f, students)
#   return students

def makeStruct(f, students):

  read = f.read().replace('\r', '')
  lines = read.split('\n')

  lines = lines[1:]
  for line in lines:
    fields = line.split('\t')
    alums = fields[1]
    sName = fields[2]
    brokenUp = alums.split('---')
    i = 0
    names = []
    for i in range(len(brokenUp)-1):
      if i % 2 == 0:
        name = brokenUp[i]
        if comma in name:
          name = name[name.rindex(comma)+4:]
        names.append(name)
    students[sName] = names
  return students

def makeChoice(students, selections, notMatched):
  student = random.choice(students.keys())
  if not students[student]:
    return

  alum = random.choice(students[student])

  selections[student] = alum
  left = students[student]
  for l in left:
    notMatched.add(l)
  notMatched.remove(alum)

  del students[student]
  for k,v in students.iteritems():
    alums = list(students[k])
    if alum in alums:
      alums.remove(alum)
    students[k] = alums

def makeChoices(ss):
  students = dict(ss)
  selections = dict()
  notMatched = set()
  i = 0
  while students and i < 100:
    makeChoice(students, selections, notMatched)
    i += 1
  return len(students), selections, notMatched

sss = makeStudents()
# sss = makeAlums(ss)

for i in range(1000):
  ll, selections, notMatched = makeChoices(sss)
  if ll == 0:
    for k,v in selections.iteritems():
      print k,',',v
    for n in notMatched:
      print n
