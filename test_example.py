from playwright.sync_api import sync_playwright
import openpyxl

# This script collects all project names/links from the Affinius Capital website's Portfolio page

def get_affinius_projects():
    with sync_playwright() as p:
        browser = p.webkit.launch()
        page = browser.new_page()
        # Go to the Portfolio page where projects are likely listed
        page.goto("https://affiniuscapital.com/portfolio/")
        page.wait_for_load_state("networkidle", timeout=60000)
        # Try to select project elements (adjust selector as needed)
        projects = page.query_selector_all('a, .project, .portfolio-item, .card, .project-title')
        project_list = []
        for proj in projects:
            text = proj.inner_text().strip()
            href = proj.get_attribute('href')
            if text and href and 'portfolio' in href:
                project_list.append({'name': text, 'url': href})
        browser.close()
        return project_list



def get_affinius_all_projects():
    with sync_playwright() as p:
        browser = p.webkit.launch()
        page = browser.new_page()
        page.goto("https://affiniuscapital.com/portfolio/")
        page.wait_for_load_state("networkidle", timeout=60000)
        # Ensure the 'ALL' tab is selected (should be by default, but click if needed)
        try:
            all_tab = page.query_selector("text=ALL")
            if all_tab:
                all_tab.click()
                page.wait_for_timeout(2000)
        except Exception:
            pass
        # Find all project name and location pairs by traversing the DOM as on the site
        projects = []
        # The pattern is: location (in all caps), then project name (h4), then type (e.g. Multifamily)
        # Find all h4s (project names), then look for the previous sibling for location
        h4s = page.query_selector_all('h4')
        for h4 in h4s:
            name = h4.inner_text().strip()
            location = 'Unknown'
            proj_type = ''
            # Get location (previous sibling)
            try:
                location_el = h4.evaluate_handle('el => el.previousElementSibling')
                if location_el:
                    loc_text = location_el.inner_text().strip()
                    # Heuristic: location is usually all uppercase and not empty
                    if loc_text and loc_text.isupper():
                        location = loc_text
            except Exception:
                pass
            # Get type (next sibling)
            try:
                type_el = h4.evaluate_handle('el => el.nextElementSibling')
                if type_el:
                    type_text = type_el.inner_text().strip()
                    proj_type = type_text
            except Exception:
                pass
            if name:
                projects.append({'name': name, 'location': location, 'type': proj_type})
        browser.close()
        return projects

def save_projects_to_excel(projects, filename=None):
    if filename is None:
        filename = "affinius_projects.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Projects"
    # Dynamically get all keys for header
    if projects:
        headers = list(projects[0].keys())
        ws.append([header.replace('_', ' ').title() for header in headers])
        # Sort projects by 'type' if present
        if 'type' in headers:
            projects = sorted(projects, key=lambda x: x.get('type', '').lower())
        for proj in projects:
            ws.append([proj.get(header, '') for header in headers])
    wb.save(filename)

def filter_multifamily_projects(projects, keywords=None):
    if keywords is None:
        keywords = ['multifamily']
    multifamily_projects = []
    for proj in projects:
        proj_type = proj.get('type', '').lower()
        if any(keyword in proj_type for keyword in keywords):
            multifamily_projects.append(proj)
    return multifamily_projects

def filter_industrial_logistics_projects(projects, keywords=None):
    if keywords is None:
        keywords = ['industrial', 'logistics']
    filtered_projects = []
    for proj in projects:
        proj_type = proj.get('type', '').lower()
        if any(keyword in proj_type for keyword in keywords):
            filtered_projects.append(proj)
    return filtered_projects



def deduplicate_projects(projects):
    seen = set()
    unique_projects = []
    for proj in projects:
        key = (proj['name'].lower(), proj['location'].lower())
        if key not in seen:
            seen.add(key)
            unique_projects.append(proj)
    return unique_projects

if __name__ == "__main__":
    all_projects = get_affinius_all_projects()
    all_projects = deduplicate_projects(all_projects)
    if all_projects:
        print("\nAffinius Capital Representative Projects (Name, Location):\n" + "-"*60)
        for proj in all_projects:
            print(f"- {proj['name']}, {proj['location']}")
        save_projects_to_excel(all_projects)
        print("\nProjects have also been saved to affinius_projects.xlsx.")

        # Filter and display Multifamily projects
        multifamily_projects = filter_multifamily_projects(all_projects)
        if multifamily_projects:
            print("\nOut of these, the Multifamily projects are:")
            for proj in multifamily_projects:
                print(f"- {proj['name']}, {proj['location']}")
        else:
            print("\nNo Multifamily projects found.")

        # Filter and display Industrial/Logistics projects
        industrial_projects = filter_industrial_logistics_projects(all_projects)
        if industrial_projects:
            print("\nOut of these, the Industrial/Logistics projects are:")
            for proj in industrial_projects:
                print(f"- {proj['name']}, {proj['location']}")
        else:
            print("\nNo Industrial/Logistics projects found.")
    else:
        print("No representative projects found or selector needs adjustment.")