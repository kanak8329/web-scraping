# import asyncio
# from playwright.async_api import async_playwright
# import pandas as pd
#
# async def scrape_sample(n=300):
#
#     base_listing = "https://www.topuniversities.com/world-university-rankings?items_per_page=150&page=2"
#
#     results = []
#
#     async with async_playwright() as pw:
#         browser = await pw.chromium.launch(headless=True)
#         page = await browser.new_page()
#
#         # Go to listing page
#         await page.goto(base_listing, timeout=60000)
#         # scroll to load links
#         for _ in range(20):
#             await page.mouse.wheel(0, 1000)
#             await page.wait_for_timeout(500)
#
#         # collect profile links
#         anchors = await page.query_selector_all("a[href*='/universities/']")
#         uni_urls = []
#         for a in anchors:
#             href = await a.get_attribute("href")
#             if href and "/universities/" in href:
#                 full = href if href.startswith("http") else ("https://www.topuniversities.com" + href)
#                 uni_urls.append(full)
#         uni_urls = list(dict.fromkeys(uni_urls))[:n]
#         print("Universities to scrape:", uni_urls)
#
#         for url in uni_urls:
#             print("Scraping:", url)
#             await page.goto(url, timeout=60000)
#             await page.wait_for_timeout(2000)
#
#             # Try click Students & Staff tab if exists
#             try:
#                 tab = await page.query_selector("a[href*='students-staff']")
#                 if tab:
#                     await tab.click()
#                     await page.wait_for_timeout(1500)
#             except:
#                 pass
#
#             # helper to get text next to a label
#             async def get_by_label(text_label):
#                 # find element whose text contains label, then get sibling or following
#                 el = await page.query_selector(f"text={text_label}")
#                 if el:
#                     # try next sibling
#                     sib = await el.evaluate_handle("e => e.nextElementSibling")
#                     if sib:
#                         try:
#                             return await sib.text_content()
#                         except:
#                             pass
#                 return None
#
#             uni_name = (await page.text_content("h1")) or url
#
#             total_students = await get_by_label("Total students")
#             intl_students = await get_by_label("International students")
#             faculty_staff = await get_by_label("Total faculty staff")
#             # For staff percentages, two labels: “Domestic staff” and “Int'l staff”
#             dom_staff_pct = await get_by_label("Domestic staff")
#             int_staff_pct = await get_by_label("Int'l staff")
#             # If one missing, try alternate label (with unicode apostrophe)
#             if int_staff_pct is None:
#                 int_staff_pct = await get_by_label("Int’ l staff")  # alternate
#             if dom_staff_pct is None:
#                 dom_staff_pct = await get_by_label("Domestic staff")
#
#             results.append({
#                 "University": uni_name.strip(),
#                 "Total Students": (total_students or "").strip(),
#                 "International Students": (intl_students or "").strip(),
#                 "Total Faculty Staff": (faculty_staff or "").strip(),
#                 "Domestic Staff %": (dom_staff_pct or "").strip(),
#                 "Int'l Staff %": (int_staff_pct or "").strip()
#             })
#
#         await browser.close()
#
#     df = pd.DataFrame(results)
#     df.to_excel("qs_table_sample03.xlsx", index=False)
#     print("Done. Sample saved to qs_table_sample03.xlsx")
#
# if __name__ == "__main__":
#     asyncio.run(scrape_sample(300))


# import asyncio
# from playwright.async_api import async_playwright
# import pandas as pd
#
# async def scrape_sample(start=151, end=300):
#
#     base_listing = "https://www.topuniversities.com/world-university-rankings?page=1&items_per_page=150&sort_by=rank&order_by=asc"
#
#     results = []
#
#     async with async_playwright() as pw:
#         browser = await pw.chromium.launch(headless=True)
#         page = await browser.new_page()
#
#         # Go to listing page
#         await page.goto(base_listing, timeout=60000)
#         # scroll to load links
#         for _ in range(20):
#             await page.mouse.wheel(0, 1000)
#             await page.wait_for_timeout(500)
#
#         # collect profile links
#         anchors = await page.query_selector_all("a[href*='/universities/']")
#         uni_urls = []
#         for a in anchors:
#             href = await a.get_attribute("href")
#             if href and "/universities/" in href:
#                 full = href if href.startswith("http") else ("https://www.topuniversities.com" + href)
#                 uni_urls.append(full)
#
#         # Deduplicate and slice from start to end
#         uni_urls = list(dict.fromkeys(uni_urls))[start-1:end]
#         print("Universities to scrape:", len(uni_urls))
#
#         for url in uni_urls:
#             print("Scraping:", url)
#             await page.goto(url, timeout=60000)
#             await page.wait_for_timeout(2000)
#
#             # Try click Students & Staff tab if exists
#             try:
#                 tab = await page.query_selector("a[href*='students-staff']")
#                 if tab:
#                     await tab.click()
#                     await page.wait_for_timeout(1500)
#             except:
#                 pass
#
#             # helper to get text next to a label
#             async def get_by_label(text_label):
#                 el = await page.query_selector(f"text={text_label}")
#                 if el:
#                     sib = await el.evaluate_handle("e => e.nextElementSibling")
#                     if sib:
#                         try:
#                             return await sib.text_content()
#                         except:
#                             pass
#                 return None
#
#             uni_name = (await page.text_content("h1")) or url
#
#             total_students = await get_by_label("Total students")
#             intl_students = await get_by_label("International students")
#             faculty_staff = await get_by_label("Total faculty staff")
#             dom_staff_pct = await get_by_label("Domestic staff")
#             int_staff_pct = await get_by_label("Int'l staff")
#             if int_staff_pct is None:
#                 int_staff_pct = await get_by_label("Int’ l staff")
#
#             results.append({
#                 "University": uni_name.strip(),
#                 "Total Students": (total_students or "").strip(),
#                 "International Students": (intl_students or "").strip(),
#                 "Total Faculty Staff": (faculty_staff or "").strip(),
#                 "Domestic Staff %": (dom_staff_pct or "").strip(),
#                 "Int'l Staff %": (int_staff_pct or "").strip()
#             })
#
#         await browser.close()
#
#     df = pd.DataFrame(results)
#     df.to_excel("qs_table_sample02.xlsx", index=False)
#     print("Done. Sample saved to qs_table_sample.xlsx")
#
# if __name__ == "__main__":
#     asyncio.run(scrape_sample(151, 300))


import asyncio
import math
from playwright.async_api import async_playwright, TimeoutError as PlaywrightTimeoutError
import pandas as pd
import time
import random


async def scrape_qs(total=1503, per_page=150, headless=True):
    base = "https://www.topuniversities.com/world-university-rankings"
    num_pages = math.ceil(total / per_page)  # e.g. 1503 / 150 -> 11 pages (0..10)
    uni_urls = []
    results = []

    async with async_playwright() as pw:
        browser = await pw.chromium.launch(headless=headless)
        # use a context so we can open new tabs without losing listing page
        context = await browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120 Safari/537.36"
        )
        page = await context.new_page()

        # 1) collect profile URLs from each listing page
        for page_idx in range(num_pages):
            listing_url = f"{base}?items_per_page={per_page}&page={page_idx}"
            print(f"[Listing] Loading page {page_idx + 1}/{num_pages}: {listing_url}")
            try:
                await page.goto(listing_url, timeout=60000)
            except Exception as e:
                print("  goto failed:", e)
                continue

            # wait for at least one university link to appear
            try:
                await page.wait_for_selector("a[href*='/universities/']", timeout=20000)
            except PlaywrightTimeoutError:
                print("  no anchors found on listing page", page_idx)

            # scroll a bit to ensure lazy loaded items appear
            for _ in range(8):
                await page.mouse.wheel(0, 1000)
                await page.wait_for_timeout(300)

            # use page.evaluate to collect hrefs (normalized to absolute)
            urls = await page.evaluate("""
                () => {
                    const all = Array.from(document.querySelectorAll("a[href*='/universities/']"));
                    // prefer anchors inside listing rows if possible (heuristic)
                    const filtered = all.filter(a => a.closest('.views-row') || a.closest('[class*="ranking"]') || a.closest('.qs-rankings') || a.closest('.uni-listing'));
                    const source = filtered.length ? filtered : all;
                    const hrefs = source.map(a => {
                        const h = a.getAttribute('href');
                        if (!h) return null;
                        if (h.startsWith('http')) return h;
                        if (h.startsWith('/')) return location.origin + h;
                        // relative path fallback
                        return new URL(h, location.href).href;
                    }).filter(Boolean);
                    // dedupe preserving order
                    return Array.from(new Set(hrefs));
                }
            """)
            print(f"  found {len(urls)} anchors (listing page {page_idx + 1})")
            uni_urls.extend(urls)

            # small polite random wait
            await page.wait_for_timeout(500 + random.randint(200, 800))

        # dedupe across pages and limit to total
        uni_urls = list(dict.fromkeys(uni_urls))[:total]
        print("Total unique university URLs collected:", len(uni_urls))
        if len(uni_urls) < total:
            print(
                "Warning: collected fewer URLs than requested. You may need to increase scrolling/waiting or check page structure.")

        # 2) Visit each profile page and scrape details
        for idx, url in enumerate(uni_urls, start=1):
            print(f"[{idx}/{len(uni_urls)}] Scraping: {url}")
            profile = await context.new_page()
            try:
                await profile.goto(url, timeout=60000)
                await profile.wait_for_load_state("networkidle")
                await profile.wait_for_timeout(1000)  # small wait for dynamic content

                # try to click Students & Staff tab if it exists
                try:
                    # try a couple of selectors: direct href, visible text, or aria
                    tab = await profile.query_selector("a[href*='students-staff']")
                    if not tab:
                        tab = await profile.query_selector("text='Students & staff'")
                    if not tab:
                        tab = await profile.query_selector("text=Students & staff")
                    if tab:
                        await tab.click()
                        await profile.wait_for_timeout(900)
                except Exception:
                    pass

                async def get_by_label(p, label):
                    el = await p.query_selector(f"text={label}")
                    if not el:
                        return None
                    try:
                        sib = await el.evaluate_handle("e => e.nextElementSibling")
                        if not sib:
                            return None
                        txt = await sib.text_content()
                        return txt.strip() if txt else None
                    except Exception:
                        return None

                uni_name = (await profile.text_content("h1")) or url
                total_students = await get_by_label(profile, "Total students")
                intl_students = await get_by_label(profile, "International students")
                faculty_staff = await get_by_label(profile, "Total faculty staff")
                dom_staff_pct = await get_by_label(profile, "Domestic staff")
                int_staff_pct = await get_by_label(profile, "Int'l staff") or await get_by_label(profile,
                                                                                                 "Int’ l staff")

                results.append({
                    "University": uni_name.strip() if uni_name else url,
                    "Total Students": (total_students or "").strip(),
                    "International Students": (intl_students or "").strip(),
                    "Total Faculty Staff": (faculty_staff or "").strip(),
                    "Domestic Staff %": (dom_staff_pct or "").strip(),
                    "Int'l Staff %": (int_staff_pct or "").strip()
                })

            except Exception as e:
                print("  Failed to scrape", url, "-", e)
            finally:
                await profile.close()
                # polite pause every few requests
                if idx % 10 == 0:
                    await asyncio.sleep(1 + random.random() * 2)

        await browser.close()

    # 3) save
    df = pd.DataFrame(results)
    df.to_excel(f"qs_table_{len(results)}.xlsx", index=False)
    print("Done. Saved to qs_table_{}.xlsx".format(len(results)))


if __name__ == "__main__":
    asyncio.run(scrape_qs(total=1503, per_page=150, headless=True))
