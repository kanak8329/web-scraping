from pybliometrics.scopus import CitationOverview

# Example: Citation overview for a Scopus document
eid = '2-s2.0-85068268017'  # replace with your paper’s EID
co = CitationOverview([eid], start=2010, end=2025, refresh=True)

print("Total citations:", co.citations_count)
print("Citations per year:", co.cc)
