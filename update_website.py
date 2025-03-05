import pandas as pd
from datetime import datetime

def html_txt_and_link_if_any(row, item, item_link, with_comma_and_space=False):
    txt = ''
    if not pd.isna(row[item]):
        if pd.isna(row[item_link]): txt += str(row[item])
        else: txt += "<a href=\"" + row[item_link] + "\" target=\"_blank\">" + row[item] + "</a>"
        if with_comma_and_space: txt += ", "
    return txt

def new_symbol_if_new(row):
    txt = ""
    if row["new"]=="yes": txt+= "&#127381; "
    return txt

def clamped_html_link_if_any(row, item_link, name, extra_text=None):
    txt = ""
    if not pd.isna(row[item_link]):
        txt += "[<a href=\"" + row[item_link] + "\" target=\"_blank\">" + name + "</a>"
        if extra_text!=None: txt+=extra_text
        txt += "] "
    return txt

def pictured_if_pictured(row):
    txt = ""
    if row["pictured"]=="yes": txt+= " (pictured &#8594;)"
    return txt

# Read the header HTML file
with open('website_header.html', 'r', encoding='utf-8') as file1:
    header_html = file1.read()
output_html = header_html

# --------------------- PUBLICATIONS ------------------------#

data = pd.read_excel(r"info.xlsx", sheet_name="publications")
output_html += "<h2>Publications</h2>\n"
for index,row in data.iterrows():
    txt =  "<img src=" + row["image"] + " alt=" + row["image"] + " class=\"projecticon\">\n"
    txt += "<p class=\"small_2\">" + row["type"] + "</p>\n"
    txt += "<h4 style=\"text-decoration:none\" class=\"tight\">"
    txt += new_symbol_if_new(row)
    txt += "<u>" + row["title"] + "</u></h4>\n"
    txt += "<p style=\"margin-top:0em;\">\n"
    txt += row["journal_conference_name_year"]
    if not pd.isna(row["award"]):  txt += ", &#127942; " + row["award"] 
    txt += "<br>\n"
    txt += row["author_list"] + "<br>\n"
    txt += clamped_html_link_if_any(row, "pdf",   "PDF")
    txt += clamped_html_link_if_any(row, "preprint_pdf", "Preprint PDF")
    txt += clamped_html_link_if_any(row, "video", "Video")
    txt += clamped_html_link_if_any(row, "doi",   "DOI")
    txt += clamped_html_link_if_any(row, "project_page", "Project page")
    txt += clamped_html_link_if_any(row, "github", "Code")
    if not pd.isna(row["github_stars"]):
        txt += " <b>&#9734; "+ str(row["github_stars"]) +" stars</b> on GitHub"
    txt += clamped_html_link_if_any(row, "abstract_pdf", "Abstract PDF")
    txt += clamped_html_link_if_any(row, "abstract_pdf_jp", "Abstract PDF (JP)")    
    txt+="<br clear=\"left\"><br></p>\n\n"
    output_html+=txt

# --------------------- OTHER WRITING ------------------------#

data = pd.read_excel(r"info.xlsx", sheet_name="other_writing")
output_html += "<p><i>Other (non-peer-reviewed) writing:</i></p>\n"
for index,row in data.iterrows():
    txt = "<p class=\"small_2\">" + row["type"] + "</p>\n"
    txt += "<h4 style=\"text-decoration:none\" class=\"tight\">"
    txt += new_symbol_if_new(row)
    txt += "<u>" + row["title"] + "</u></h4>\n"
    txt += "<p style=\"margin-top:0em;\">\n"
    if not pd.isna(row["english_title"]): txt += "English title: " + row["english_title"] + "\n"
    txt += row["publisher"] + ", "
    if isinstance(row['date'], datetime): txt += str(row["date"].date())
    else: txt += row["date"]
    txt += " " + clamped_html_link_if_any(row, "link",   "Link")
    txt+="<br></p>\n\n"
    output_html+=txt

# --------------------- GRANTS AND AWARDS ------------------------#

data = pd.read_excel(r"info.xlsx", sheet_name="grants_and_awards")
output_html += "<h2>Grants & Awards</h2>\n<ul>"
for index,row in data.iterrows():
    txt =  '<li>'
    txt += new_symbol_if_new(row)
    txt += html_txt_and_link_if_any(row, "name", "name_link", with_comma_and_space=True)
    if not pd.isna(row["agency"]): 
        txt += html_txt_and_link_if_any(row, "agency", "agency_link", with_comma_and_space=True)
    txt += str(row["year"])
    txt += "</li>\n"
    output_html+=txt
output_html += "</ul>\n\n"

# --------------------- PRESENTATIONS ------------------------#

data = pd.read_excel(r"info.xlsx", sheet_name="presentations")
output_html += "<h2>Presentations</h2>\n<ul>"
for index,row in data.iterrows():
    txt =  '<li>'
    txt += new_symbol_if_new(row)
    txt += row["type"] + ", "
    if not pd.isna(row["event"]): 
        txt += html_txt_and_link_if_any(row, "event", "event_page", with_comma_and_space=True)
    if not pd.isna(row["organization"]): 
        txt += html_txt_and_link_if_any(row, "organization", "organization_page", with_comma_and_space=True)
    txt += row["location"] + ", "
    if isinstance(row['date'], datetime): txt += str(row["date"].date())
    else: txt += row["date"]
    txt += "</li>\n"
    output_html+=txt
output_html += "</ul>\n\n"

# --------------------- EXHIBITIONS ------------------------#

data = pd.read_excel(r"info.xlsx", sheet_name="exhibitions")
output_html += "<h2>Exhibitions</h2>\n"
output_html += "<a href=swirl.html target=_blank><img src=branch3d/swirled_img.jpg class=righticon_toppad></a>"
output_html += "<ul>"
for index,row in data.iterrows():
    txt =  '<li>'
    txt += new_symbol_if_new(row)
    txt += html_txt_and_link_if_any(row, "name", "name_link", with_comma_and_space=True)
    txt += html_txt_and_link_if_any(row, "venue", "venue_link", with_comma_and_space=True)
    txt += row["location"] + ', '
    txt += str(row["year"])
    txt += pictured_if_pictured(row)
    txt += "</li>\n"
    output_html+=txt
output_html += "</ul>\n\n"

# --------------------- DESIGN WORK ------------------------#

data = pd.read_excel(r"info.xlsx", sheet_name="design_work")
output_html += "<h2>Design Work</h2>\n"
output_html += "<a href=https://www.poolarch.ch/projekte/2015/campus-biel.html&refPage=projekte target=_blank><img src=design_work/biel_img.jpg alt=Campus Biel/Bienne class=righticon></a>"
output_html += "<ul>"
for index,row in data.iterrows():
    txt =  '<li>'
    txt += new_symbol_if_new(row)
    txt += html_txt_and_link_if_any(row, "project", "project_link", with_comma_and_space=True)
    txt += html_txt_and_link_if_any(row, "office", "office_link", with_comma_and_space=True)
    txt += str(row["year"])
    txt += pictured_if_pictured(row)
    txt += "</li>\n"
    output_html+=txt
output_html += "</ul>\n\n"

# --------------------- MEDIA ------------------------#

data = pd.read_excel(r"info.xlsx", sheet_name="media")
output_html += "<h2>Media</h2>\n"
output_html += "<a href=https://www.lemonde.fr/sciences/article/2020/11/11/infographie-automatiser-l-art-du-joint-a-la-japonaise-sans-clou-ni-colle_6059379_1650684.html target=_blank><img src=media/lemon_img.jpg class=righticon></a>"
output_html += "<ul>"
for index,row in data.iterrows():
    txt =  '<li>'
    txt += new_symbol_if_new(row)
    txt += html_txt_and_link_if_any(row, "publisher", "publisher_link", with_comma_and_space=True)
    if isinstance(row['date'], datetime): txt += str(row["date"].date())
    else: txt += row["date"]
    txt += pictured_if_pictured(row)
    txt += "</li>\n"
    output_html+=txt
output_html += "</ul>\n\n"

# --------------------- SERVICE ------------------------#

data = pd.read_excel(r"info.xlsx", sheet_name="service")
output_html += "<h2>Service</h2>\n<ul>"
for index,row in data.iterrows():
    txt =  '<li>'
    txt += new_symbol_if_new(row)
    txt += html_txt_and_link_if_any(row, "role", "role_link", with_comma_and_space=False)
    if not pd.isna(row["organization"]): txt += ", "
    txt += html_txt_and_link_if_any(row, "organization", "org_link", with_comma_and_space=False)
    if not pd.isna(row['year']): txt += ", " + str(row["year"])
    txt += "</li>\n"
    output_html+=txt
output_html += "</ul>\n\n"

# ---- FOOTER -----#

output_html += "<br></body>\n</html>"

# --------------------- SAVE ------------------------#

with open('home.html', 'w',  encoding='utf-8') as output_file:
    output_file.write(output_html)
    print("Saved", 'home.html')

with open('about.html', 'w',  encoding='utf-8') as output_file:
    output_file.write(output_html)
    print("Saved", 'about.html')

with open('index.html', 'w',  encoding='utf-8') as output_file:
    output_file.write(output_html)
    print("Saved", 'index.html')

