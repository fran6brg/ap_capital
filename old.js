


        
  console.log("\nscrapping company " + row_index + " : " + fail.site);
  try {
      // catch if refused connection
      await page.goto(fail.site, { waitUntil: 'domcontentloaded' });
      await page.setViewport({
          width: 1200,
          height: 800
      });
      // await page.waitForTimeout(5000);
      try {
          try {
              console.log(chalk.blue("    get hrefs inside page: " + fail.site));
              await autoScroll(page);
              const hrefs = await page.evaluate(
                  () => Array.from(
                      document.querySelectorAll('a[href]'),
                      a => a.getAttribute('href')
                  )
              );
              // console.log(hrefs);

              const cell_linkedin_link = worksheet.getCell('C' + row_index);
              const cell_taille = worksheet.getCell('D' + row_index);
              const cell_nb_employees = worksheet.getCell('E' + row_index);

              let found = false;
              let link_completed = "";
              for (var i = 0; i < hrefs.length; i++) {
                  if (hrefs[i].includes(".linkedin.com/")
                      /*hrefs[i].startsWith("https://www.linkedin.com/")
                      || hrefs[i].startsWith("https://linkedin.com/")
                      || hrefs[i].startsWith("https://fr.linkedin.com/")
                      || hrefs[i].startsWith("//www.linkedin.com/")
                      || hrefs[i].startsWith("//fr.linkedin.com/")*/) {
                          
                      console.log(chalk.green("    found href linkedin: " + hrefs[i]));
                      found = true;
                      try {
                          link_completed = hrefs[i];
                          cell_linkedin_link.value = link_completed;
                          if (!link_completed.includes("/company/")) {
                              cell_taille.value = "not a company link";
                              cell_nb_employees.value = "not a company link";
                              workbook.xlsx.writeFile('linkedin_v2.xlsx');
                              console.log(chalk.red("    not a company link"));
                              continue ;
                          } else if (link_completed.endsWith("/about/")) {
                              ; // ok perfect
                          } else if (link_completed.endsWith("/about")) {
                              ; // ok perfect
                          } else if (link_completed.endsWith("/")) {
                              let pos = link_completed.split('/', 5).join('/').length;
                              link_completed = link_completed.substr(0, pos);
                              pos = link_completed.split('?', 1).join('?').length;
                              link_completed = link_completed.substr(0, pos);
                              link_completed = link_completed.concat("/about/");

                          } else { // does't end with '/'
                              let pos = link_completed.split('/', 5).join('/').length;
                              link_completed = link_completed.substr(0, pos);
                              pos = link_completed.split('?', 1).join('?').length;
                              link_completed = link_completed.substr(0, pos);
                              link_completed = link_completed.concat("/about/");
                          }
                          cell_linkedin_link.value = link_completed;
                          console.log("    link_completed : " + link_completed);
                      } catch (error) {
                          console.log(chalk.red("    error link_completed: " + error));
                      }

                      console.log("    browsing linkedin : " + link_completed + " ...");
                      await page.goto(link_completed, { waitUntil: 'domcontentloaded' });
                      // await page.waitForTimeout(3000);

                      // get elements innerText
                      try {
                          
                          let selector_taille = 'body > div.application-outlet > div.authentication-outlet > div > div.org-organization-page__container > div.org-grid.mt5 > div.org-grid__core-rail--wide.mb6 > div.org-grid__core-rail--no-margin-left > div:nth-child(1) > section > dl > dd.org-about-company-module__company-size-definition-text.t-14.t-black--light.mb1.fl';
                          
                          // if already present on page
                          try {
                              await page.waitForSelector(selector_taille)
                          } catch (error) {
                              try {
                                  selector_taille = 'body > div.application-outlet > div.authentication-outlet > div > div.org-organization-page__container > div.org-grid.mt5 > div.org-grid__core-rail--wide.mb6 > div.org-grid__core-rail--no-margin-left > div:nth-child(1) > section > dl > dd.org-about-company-module__company-size-definition-text.t-14.t-black--light.mb5'
                                  console.log(chalk.red("    change taille selector"));
                                  await page.waitForSelector(selector_taille)
                                  const element_taille = await page.$(selector_taille)
                                  taille = await page.evaluate(el => el.textContent, element_taille)
                                  taille = taille.trim();
                                  console.log(chalk.green("    taille: " + taille));
                                  cell_taille.value = taille;
                                  cell_nb_employees.value = "na";
                                  workbook.xlsx.writeFile('linkedin_v2.xlsx');
                                  continue ;
                              } catch (error) {
                                  console.log(chalk.red("    error waitForSelector: " + error));
                                  cell_taille.value = "broke url, can't fetch /about";
                                  cell_nb_employees.value = "broke url, can't fetch /about";
                                  workbook.xlsx.writeFile('linkedin_v2.xlsx');
                                  continue ;
                                  
                              }
                          }

                          console.log("    searching for selector_taille");
                          // let selector_taille = 'body > div.application-outlet > div.authentication-outlet > div > div.org-organization-page__container > div.org-grid.mt5 > div.org-grid__core-rail--wide.mb6 > div.org-grid__core-rail--no-margin-left > div:nth-child(1) > section > dl > dd.org-about-company-module__company-size-definition-text.t-14.t-black--light.mb1.fl';
                          let taille = "";
                          try {
                              const element_taille = await page.$(selector_taille)
                              taille = await page.evaluate(el => el.textContent, element_taille)
                              taille = taille.trim();
                              console.log(chalk.green("    taille: " + taille));
                              cell_taille.value = taille;
                          } catch (error) {
                              console.log(chalk.red("    error taille: " + error));
                          }

                          console.log("    searching for selector_nb_employees");
                          let selector_nb_employees = 'body > div.application-outlet > div.authentication-outlet > div > div.org-organization-page__container > div.org-grid.mt5 > div.org-grid__core-rail--wide.mb6 > div.org-grid__core-rail--no-margin-left > div:nth-child(1) > section > dl > dd.org-page-details__employees-on-linkedin-count.t-14.t-black--light.mb5';
                          // await page.waitForSelector(selector_nb_employees)
                          let nb_employees = "";
                          try {
                              const element_nb_employees = await page.$(selector_nb_employees)
                              nb_employees = await page.evaluate(el => el.textContent, element_nb_employees)
                              const searchTerm = 'Inclut des membres';
                              const indexOfFirst = nb_employees.indexOf(searchTerm);
                              if (indexOfFirst != -1) {
                                  nb_employees = nb_employees.substring(0, indexOfFirst - 1)
                              }
                              nb_employees = nb_employees.trim();
                              console.log(chalk.green("    nb employee: " + nb_employees));
                              cell_nb_employees.value = nb_employees;
                          } catch (error) {
                              console.log(chalk.red("    error nb_employees: " + error));
                          }
                      } catch (error) {
                          console.log(chalk.red("    error getting elements: " + error));
                      }
                      try {
                          workbook.xlsx.writeFile('linkedin_v2.xlsx');
                      } catch (error) {
                          console.log(chalk.red("    error writeFile: " + error));
                      }
                      break;
                  }
              }
              if (found === false) {
                  const cell_linkedin_link = worksheet.getCell('C' + row_index);
                  cell_linkedin_link.value = "no linkedin link on site";
                  const cell_taille = worksheet.getCell('D' + row_index);
                  cell_taille.value = "no linkedin link on site";
                  const cell_nb_employees = worksheet.getCell('E' + row_index);
                  cell_nb_employees.value = "no linkedin link on site";
                  workbook.xlsx.writeFile('linkedin_v2.xlsx');
                  console.log(chalk.red("    no linkedin found"));
              }
              row_index += 1;

          } catch (error) {
              const cell_linkedin_link = worksheet.getCell('C' + row_index);
              // cell_linkedin_link.value = "unhandled error";
              const cell_taille = worksheet.getCell('D' + row_index);
              cell_taille.value = "unhandled error";
              const cell_nb_employees = worksheet.getCell('E' + row_index);
              cell_nb_employees.value = "unhandled error";
              workbook.xlsx.writeFile('linkedin_v2.xlsx');
              console.log(chalk.red("    error while getting hrefs: " + error));
              continue;
          }
      } catch (error) {
          const cell_linkedin_link = worksheet.getCell('C' + row_index);
          cell_linkedin_link.value = "unhandled error";
          const cell_taille = worksheet.getCell('D' + row_index);
          cell_taille.value = "unhandled error";
          const cell_nb_employees = worksheet.getCell('E' + row_index);
          cell_nb_employees.value = "unhandled error";
          workbook.xlsx.writeFile('linkedin_v2.xlsx');
          console.log(chalk.red("    error while remaining: " + error));
          continue;
      }
  }
  catch (error) {
      const cell_linkedin_link = worksheet.getCell('C' + row_index);
      cell_linkedin_link.value = "refused connection";
      const cell_taille = worksheet.getCell('D' + row_index);
      cell_taille.value = "refused connection";
      const cell_nb_employees = worksheet.getCell('E' + row_index);
      cell_nb_employees.value = "refused connection";
      workbook.xlsx.writeFile('linkedin_v2.xlsx');
      console.log(chalk.red("    error while remaining: " + error));
      console.log(chalk.red("    refused connection when page.goto(url): " + error));
      continue;
  }
}
await browser.close();
})();