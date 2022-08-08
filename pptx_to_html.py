from lxml import etree as et
import credentials
######################################################################################
#                                                                                    #
#                                                                                    #
#                             Powerpoint to html                                     #
#                             ==================                                     #
#                                                                                    #
#    1. Export from powerpoint to svg all the slides and place in input folder       #
#    2. Any custom html you want to put into the file place in prefix.html in        # 
#           the source folder                                                         #
#    3. Run this script                                                              #
#                                                                                    #
#                                                                                    #
######################################################################################

PAGE_FILENAME_PREFIX:str="Slide"
INPUT_FOLDER_PATH:str="input/"
OUTPUT_FOLDER_PATH:str="output/"
PREFIX_FILE:str="prefix.html"

def get_svg(filenap:str)->object:
    with open(filenap,mode="r",encoding="utf-8") as infile:
        content=infile.read()
    piece=et.XML(content)
    return piece


def get_all_ids(piece:et.Element)->dict[str,str]:
    stems=["clip","img"]
    ids:dict[str,str]={}
    for st_id,stem in enumerate(stems):
        #print(f"Staring:{st_id} {stem}")
        id=0
        while id<50: # Upto 50 index seems to be enough
            xp="//*[@id='"+stem+str(id)+"']"
            refs=piece.xpath(xp)
            if len(refs)>0:
                ids[stem+str(id)]=xp
                #print(xp)
            id+=1
    return ids
            


def rewrite_ids(piece:et.Element,page_index:int)->et.Element:
    """
    
    """
    ids:dict[str,str]=get_all_ids(piece)
    prefix="page"+str(page_index)+"_"
    for id,xp in ids.items():
        refs=piece.xpath(xp)
        # Change the reference
        for ref in refs:
            ref.attrib["id"]=prefix+id

        if id.startswith("img"):
            xp="//*[name()='use' and @*='#"+id+"']"
            img_refs=piece.xpath(xp)
            #print(xp,img_refs[0].attrib['{http://www.w3.org/1999/xlink}href'])
            new_href="#"+prefix+id
            for ref in img_refs:
                ref.attrib['{http://www.w3.org/1999/xlink}href']=new_href

        elif id.startswith("clip"):
            xp="//*[name()='g' and @*='#"+id+"']"
            xp="//*[name()='g' and @clip-path='url(#"+id+")']"
            clip_refs=piece.xpath(xp)
            #print (f"Xpath used: {xp}, last element result: {clip_refs[-1].attrib['clip-path']}")
            for ref in clip_refs:
                ref.attrib["clip-path"]="url('#"+prefix+id+"')"




        else:

            raise KeyError(f"Don't know how to resolve an id like : {id}")



    # Change out the clip pages
    



def pretty_tree(root:et.Element)->str:
    return et.tostring(root, pretty_print=True).decode("utf-8")
    
def save_tree(root:et.Element,filenap:str)->None:
    with open(filenap,"w") as outfile:
        outfile.write(et.tostring(root).decode("utf-8"))


if __name__=="__main__":

    tracking_tag=input("Enter the tracking code you want to pick?\n")

    root = et.Element('html', version="5.0") # the output tree

    tracking_img=et.Element("img")
    tracking_img.attrib["src"]=credentials.TRACKING_BASE_URL+tracking_tag

    root.append(tracking_img)
    
    index=1
    while True:
        try:
            piece=get_svg(INPUT_FOLDER_PATH+PAGE_FILENAME_PREFIX+str(index)+".svg")
        except FileNotFoundError:
            break
        rewrite_ids(piece,index)

        root.append(piece)
        index+=1
    
    #print(pretty_tree(root))
    save_tree(root,OUTPUT_FOLDER_PATH+"output.html")







