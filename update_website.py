import pandas as pd
from datetime import datetime

# Read the header HTML file
with open('website_header.html', 'r', encoding='utf-8') as file1:
    header_html = file1.read()

# Read the footer HTML file
with open('website_middle.html', 'r',  encoding='utf-8') as file2:
    middle_html = file2.read()

# Read the footer HTML file
with open('website_footer.html', 'r',  encoding='utf-8') as file3:
    footer_html = file3.read()


# --------------------- PUBLICATIONS ------------------------#

publication_data = pd.read_excel(r"info.xlsx", sheet_name="publications")
publication_html_txt = "<h2>Publications</h2>\n"
for index,row in publication_data.iterrows():

    #title = pd.DataFrame(item, columns=["product_name"])
    #print()
    #print(row['title'])

    txt =  "<img src=" + row["image"] + " alt=" + row["image"] + " class=\"projecticon\">\n"
    txt += "<p class=\"small_2\">" + row["type"] + "</p>\n"
    txt += "<h4 style=\"text-decoration:none\" class=\"tight\">"
    if row["new"]=="yes": txt+= "&#127381; "
    txt += "<u>" + row["title"] + "</u></h4>\n"
    txt += "<p style=\"margin-top:0em;\">\n"
    txt += row["journal_conference_name_year"]
    if not pd.isna(row["award"]):  txt += ", &#127942; " + row["award"] 
    txt += "<br>\n"
    txt += row["author_list"] + "<br>\n"
    if not pd.isna(row["pdf"]):              txt += "[<a href=\"" + row["pdf"]              + "\" target=\"_blank\">PDF</a>] "
    if not pd.isna(row["video"]):            txt += "[<a href=\"" + row["video"]            + "\" target=\"_blank\">Video</a>] "
    if not pd.isna(row["doi"]):              txt += "[<a href=\"" + row["doi"]              + "\" target=\"_blank\">DOI</a>] "
    if not pd.isna(row["project_page"]):     txt += "[<a href=\"" + row["project_page"]     + "\" target=\"_blank\">Project page</a>] "
    if not pd.isna(row["github"]):          
        if not pd.isna(row["github_stars"]): txt += "[<a href=\"" + row["github"]           + "\" target=\"_blank\">GitHub</a> <b>&#9734; "+ row["github_stars"] +" stars</b></a>] "
        else:                                txt += "[<a href=\"" + row["github"]           + "\" target=\"_blank\">GitHub</a>] "
    if not pd.isna(row["abstract_pdf"]):     txt += "[<a href=\"" + row["abstract_pdf"]     + "\" target=\"_blank\">Abstract PDF</a>] "
    if not pd.isna(row["abstract_pdf_jp"]):  txt += "[<a href=\"" + row["abstract_pdf_jp"]  + "\" target=\"_blank\">Abstract PDF (JP)</a>]"
    
    txt+="<br clear=\"left\"><br></p>\n\n"

    publication_html_txt+=txt


# --------------------- PRESENTATIONS ------------------------#

presentation_data = pd.read_excel(r"info.xlsx", sheet_name="presentations")
presentation_html_txt = "<h2>Presentations</h2>\n<ul>"
for index,row in presentation_data.iterrows():
    txt =  '<li>'
    if row["new"]=="yes": txt+= "&#127381; "
    txt += row["type"] + ", "
    if not pd.isna(row["event"]): 
        if not pd.isna(row["event_page"]): txt += "<a href=\"" + row["event_page"] + "\" target=\"_blank\">" + row["event"] + "</a>, "
        else: txt += row["event"] + ", "
    if not pd.isna(row["organization"]): 
        if not pd.isna(row["organization_page"]): txt += "<a href=\"" + row["organization_page"] + "\" target=\"_blank\">" + row["organization"] + "</a>, "
        else: txt += row["organization"] + ", "
    txt += row["location"] + ", "
    if isinstance(row['date'], datetime): txt += str(row["date"].date())
    else: txt += row["date"]
    txt += "</li>\n"

    presentation_html_txt+=txt

presentation_html_txt += "</ul>\n\n"





# --------------------- COMBINE and SAVE ------------------------#

# Combine everything
output_html = header_html + publication_html_txt + middle_html + presentation_html_txt + footer_html

# Write the final combined content into a new HTML file
with open('home.html', 'w',  encoding='utf-8') as output_file:
    output_file.write(output_html)
    print("Saved", 'home.html')

with open('about.html', 'w',  encoding='utf-8') as output_file:
    output_file.write(output_html)
    print("Saved", 'about.html')

with open('index.html', 'w',  encoding='utf-8') as output_file:
    output_file.write(output_html)
    print("Saved", 'index.html')

