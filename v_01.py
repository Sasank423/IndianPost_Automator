        while i <= l:
            j = 0
            ref = df.loc[i, 'RPAD Barcode No ']
            ln = df.loc[i, 'Loan No']

            if str(ref) == 'nan' or not str(ref).strip():
                i += 1
                continue

            if rt == 0:
                rt = time()

            try:
                # Step 1: Enter ref no
                ip = wait.until(EC.presence_of_element_located((By.XPATH, "//textarea[@class='flex w-full rounded-md border border-input bg-background px-3 py-2 text-base ring-offset-background placeholder:text-muted-foreground focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-ring focus-visible:ring-offset-2 disabled:cursor-not-allowed disabled:opacity-50 md:text-sm min-h-32']")))
                ip.clear()
                ip.send_keys(ref)

#                 # Step 2: Read captcha
#                 try:
#                     cap = driver.find_element(
#                         By.XPATH,
#                         "//canvas[@class='captchaCanvas border border-gray-300 rounded "
#                         "focus:ring-2 focus:ring-[#C62829] focus:border-[#C62829] focus:outline-none']"
#                     ).get_attribute('aria-label')
#                     cap = cap.lstrip('CAPTCHA security verification image. Text content: ')
#                     cap = ''.join(x.split()[-1] for x in cap.split(','))
#                 except Exception:
#                     print(f"[{ref}] Captcha read failed, refreshing page...")
#                     driver.get("https://www.indiapost.gov.in")
#                     c += 1
#                     if c >= max_retry_per_record:
#                         print(f"[{ref}] Skipping after {max_retry_per_record} failed attempts")
#                         i += 1
#                         c = 0
#                     continue
# 
#                 # Step 3: Enter captcha
#                 captcha_input = driver.find_element(
#                     By.XPATH,
#                     "//input[@class='border border-[#D6D6D6] text-black rounded-sm p-2 w-32 focus:border-[#C62829]']"
#                 )
#                 captcha_input.clear()
#                 captcha_input.send_keys(cap)
# 
#                 # Step 4: Click search button with retry
#                 for ct in range(5):
#                     try:
#                         button = wait_.until(EC.element_to_be_clickable((
#                             By.XPATH,
#                             "//div[@class='flex items-center justify-between gap-2 bg-white captch_row mt-3']"
#                             "//button[contains(@class,'searchButton')]"
#                         )))
#                         button.click()
#                         break  # break after successful click
#                     except Exception as e:
#                         print(f"[{ref}] Click attempt {ct+1} failed: {e}")
#                         sleep(1)
#                 else:
#                     print(f"[{ref}] Could not click search button after 5 attempts")
#                     c += 1
#                     if c >= max_retry_per_record:
#                         i += 1
#                         c = 0
#                     continue
                
                driver.find_element(By.XPATH, "//button[@class='gap-2 whitespace-nowrap text-sm ring-offset-background duration-300 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-ring focus-visible:ring-offset-2 disabled:pointer-events-none [&_svg]:pointer-events-none [&_svg]:size-4 [&_svg]:shrink-0 relative overflow-hidden w-11/12 mx-auto h-10 min-w-[8rem] bg-red-600 hover:bg-red-700 text-white font-medium rounded-md py-2 px-4 flex items-center justify-center transition-colors disabled:opacity-50 disabled:cursor-not-allowed']").click()
                sleep(4)
                # Step 5: Wait for result (either success table or error message)
                try:
                    wait.until(
                        EC.presence_of_element_located(
                            (By.XPATH, "//h4[@class='text-lg font-semibold text-gray-800 mb-4']")
                        )
                    )
                except Exception:
                    print(f"[{ref}] No result detected, retrying...")
                    driver.get("https://www.indiapost.gov.in")
                    c += 1
                    if c >= max_retry_per_record:
                        i += 1
                        c = 0
                    continue

                # Step 6: Extract details if available
                try:
                    ul = driver.find_element(By.XPATH,"//ul[@class='space-y-6']")
                    details = ul.find_elements(
                        By.XPATH,
                        "//li[@class='relative pl-12']"
                    )[-1].text.split('\n')
                    
                    df.loc[i, 'Delivery Report'] = details[0]
                    df.loc[i, 'date'], df.loc[i, 'time'] = details[1].split()
                    df.loc[i, 'office'] = details[2]

                    if pdf_opt:
                        pdfs.append((
                            driver.execute_cdp_cmd(
                                'Page.printToPDF', {"printBackground": False}
                            )['data'],
                            f"{ln}.pdf"
                        ))

                    df_view.dataframe(df)

                    elapsed = str(datetime.timedelta(seconds=int(time() - rt))).split(':')
                    st.write(f"{i}) Record {ref} Completed - {elapsed[1]}:{elapsed[2]}")

                    rt = 0
                    i += 1
                    c = 0  # reset failure counter
                    sleep(2)

                except Exception as e:
                    print(f"[{ref}] Data extraction failed: {e}")
                    driver.get("https://www.indiapost.gov.in")
                    c += 1
                    if c >= max_retry_per_record:
                        i += 1
                        c = 0
                    continue

            except Exception as e:
                print(f"[{ref}] Unexpected error: {e}")
                c += 1
                if c >= max_retry_per_record:
                    i += 1
                    c = 0