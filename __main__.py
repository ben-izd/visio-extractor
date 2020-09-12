import zipfile, copy, collections, os, sys


def myquit():
    input('Press any key to exit...')
    quit()

try:
    import lxml
except ImportError:
    # try to install lxml
    print('Please run the following command in your command line (Internet connection is required)')
    print('pip install lxml\n')
    myquit()
# sys.path.append(os.path.dirname(os.path.expanduser(r'D:\projects\projects\17. story_planner\final\visio\lxml')))
# sys.path.reverse()


# from lxml import etree

# path_list = [r'D:\projects\projects\17. story_planner\final\visio\lxml']

# for i in path_list:
#     sys.path.append(os.path.dirname(os.path.expanduser(i)))
# sys.path.append('lxml.zip')

# sys.path.append('bs4.zip')

# import bs4    

from bs4 import BeautifulSoup



archive_file_name = str(__file__.split('\\')[0])



class_name:str = ''

# used to rsetore program state after some errors
if len(sys.argv) > 1:
    class_name = sys.argv[1]
else:
    class_name = input('\nType your class name: (press enter for default: "Question")\n-> ')

    if class_name == '':
        class_name = 'Question'


# input : file path string (like '/pages/page1.xml')
# output: file name without extension (like 'page1')
def get_file_name(file_path:str)->str:
    return os.path.splitext(os.path.split(file_path)[1])[0]


class extract_from_visio_file:
    def __init__(self,file_name:str):
        self.file_name = os.path.splitext(file_name)[0]

        try:
            archive = zipfile.ZipFile(file_name, 'r')

        # if its a bad zip error it means user open old visio files that python can not read
        except zipfile.BadZipFile:
            print('\nSelected file is for earlier version of visio (2003-2010) that is not supported.')
            myquit()
        except:
            print('Unknown error while opening file...')
            myquit()

        print('\nAvailable pages:\n(usually is the heaviest one)')
        
        # list all files in the pages directory except pages.xml
        eligibale_files = [i for i in archive.infolist() if os.path.split(i.filename)[0] == 'visio/pages' and i.filename != 'visio/pages/pages.xml']

        # pick the heaviest item name
        heaviest_item_name = get_file_name(sorted(eligibale_files,key=lambda x: x.file_size)[-1].filename)

        for i in eligibale_files:
            # print items
            print(f'- {get_file_name(i.filename)}\t {i.file_size} \t bytes')
        
        # ask user to pick a page
        worksheet_name = input(f'\nType visio page name: (press enter for default: "{heaviest_item_name}") \n-> ').lower()
        print('')

        # if user doesnt enter any page name -> use the heaviest page name
        if worksheet_name == '':
            worksheet_name = heaviest_item_name

        # try to read the entered page
        try:
            r = archive.read(f'visio/pages/{worksheet_name}.xml').decode()
            archive.close()
        except:
            print('Counter error while reading zip file. Probably page name is invalid.')
            os.system(f'python {archive_file_name} {class_name} {target_file_name}')
            quit()
        
        
        self.b = BeautifulSoup(r,'xml')

        

        # use try beacuse some pages does not have have any shape/connection
        try:
            self.shapes = self.b.findAll('Shape')
            self.connects = self.b.find('Connects').findAll('Connect')
        except:
            print(f'there is nothing in "{worksheet_name}" page - retry with another page')
            os.system(f'python {archive_file_name} {class_name} {target_file_name}')
            quit()

        # prepare a template for connected topics
        self.connect_pairs = collections.defaultdict(list)

        # prepare a final result template
        self.object_map = collections.defaultdict(lambda : copy.deepcopy({'story':'','link_texts':[],'ids':[]}))

        # gather two way connections -> use by attribute "FromSheet"
        for connect in self.connects:
            self.connect_pairs[int(connect['FromSheet'])].append(connect)


        # iterate over all connections
        for index,pair in self.connect_pairs.items():
            
            # declare here to get access in this block
            target_id = -1
            source_id = -1

            # whether this connection is a start to end connection
            if pair[0]['FromCell'] == 'EndX':
                target_id = int(pair[0]['ToSheet'])
                source_id = int(pair[1]['ToSheet'])

            # or is a end to start connection
            else:
                source_id = int(pair[0]['ToSheet'])
                target_id = int(pair[1]['ToSheet'])
            
            # get text of start element (story)
            story = self.get_shape_text(self.find_shape_by_id(source_id))

            # get text of end element (option)
            link_text = self.get_shape_text(self.find_shape_by_id(index))

            # create condition here to doesnt save story every time we iterate over its connections
            if self.object_map[source_id]['story'] == '':
                self.object_map[source_id]['story']= self.get_shape_text(self.find_shape_by_id(source_id))

            if self.object_map[target_id]['story'] == '':
                self.object_map[target_id]['story']= self.get_shape_text(self.find_shape_by_id(target_id)) 

            # place link text in result template
            self.object_map[source_id]['link_texts'].append(link_text)

            # place link id in result template
            self.object_map[source_id]['ids'].append(target_id)


    def find_shape_by_id(self,shape_id):

        # make sure shape id is string type
        if type(shape_id) == int:
            shape_id = str(shape_id)
        
        return self.b.find('Shape',{'ID':shape_id})

    def get_shape_text(self,shape_element):
        try:
            return shape_element.find('Text').text.strip()
        except:
            return ''

    def to_string(self):
        # make final result printable
        result:str = ''
        for index,item in self.object_map.items():
            end =',\n'
            
            if index==list(self.object_map.keys())[-1]:
                end = ''
            
            result+=f"{class_name}({index}, '{item['story']}', {item['link_texts']}, {item['ids']}){end}"
        return result

    def save_to_disk(self):
        # save result file same name as opened file but with 'txt' format
        result_file_name = self.file_name+'.txt'

        # check wether is there any file similar to result file name
        if os.path.exists(result_file_name):
            print(f'ERROR - Could not save the result file - there is already a "{result_file_name}" file in the directory')
            myquit()
        try:
            with open(result_file_name,'w') as file:
                file.write(self.to_string())
            print(f'Result file ("{result_file_name}") saved successfully.')
        except:
            print(f'Counter error while saving "{result_file_name}" to disk...')
            print('Retry or apply it on a copy of your file\n')

target_file_name :str = ''

# used to restore program state after some errors
if(len(sys.argv) > 2 ):
    target_file_name = sys.argv[2]
else:
    target_file_name = input('\nType Visio file name with extension (like "1.vsdx"):\n-> ')

# check wether entered file name is exist or not
if os.path.exists(target_file_name) == False:
    print('\nERROR - either you enter invalid file name or the file is not in the directory that holds this file')
    os.system(f'python {archive_file_name} {class_name}')
    quit()


a = extract_from_visio_file(target_file_name)
a.save_to_disk()


# make program doesnt end until user wants to
myquit()