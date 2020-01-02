# SynastryParser

**SynastryParser** is a program uses csv files that contain julian date, latitude and longitude values. If users don't have an appropriate csv file they can use [ConvertGauquelinData
](https://github.com/dildeolupbiten/ConvertGauquelinData) program to generate a csv file.

## Availability

Windows, Linux, Mac

## Dependencies

In order to run **SynastryParser**, at least [Python](https://www.python.org/)'s 3.6 version must be installed on your computer. Note that in order to use [Python](https://www.python.org/) on the command prompt, [Python](https://www.python.org/) should be added to the PATH. There is no need to install manually the libraries that are used by the program. When the program first runs, the necessary libraries will be downloaded and installed automatically.

## Usage

**1.** After downloaded the **SynastryParser** script, run it by typing the following to the console window.

```
python3 SynastryParser.py
```

**2.** A window as below should be opened in a few seconds after run the above command.

![img1](https://user-images.githubusercontent.com/29302909/71669897-48945a80-2d7f-11ea-80a7-1d5f53144672.png)

**3.** As you see above, there are two menu buttons at the top of the opened window. When **"Help"** menu button is clicked, a new menu cascade is opened which name is **"Check for Updates"**. If **"Check for Updates"** menu cascade is clicked, the program checks whether there's an update of the program. The most updated version of the program is stored in GitHub.

**4.** **"Settings"** is the another menu button that users can use. When **"Settings"** menu button is clicked, three piece of menu cascades are opened which names are **"Include"**, **"Mode"**, **"Orb Factor"**.

![img2](https://user-images.githubusercontent.com/29302909/68789679-fbd1b480-0656-11ea-9434-44c6a8e246bc.png)

**5.** If **"Include"** menu cascade is selected, a window as above is opened. User should select at least one aspect checkbutton to calculate the aspects of the synastry files. For detailed analysis of aspects through signs and planets, the checkbuttons which are on the right side of the **"Include"** panel should be selected.

**6.** If **"Mode"** menu cascade is selected, a windows as below is opened. There are totally 3 modes which names are **"Natal"**, **"Antiscia"**, **"Contra-Antiscia"**. For example if users select **"Natal"** for the first person and **"Antiscia"** for the second person, the first person's natal chart positions and second person's antiscia chart positions will be used in the calculation process. 

![img3](https://user-images.githubusercontent.com/29302909/68790409-53bceb00-0658-11ea-8527-18829a1256f4.png)

**7.** The last menu cascade is for changing the **"Orb Factor"**. The default orb factors can be seen below.

![img4](https://user-images.githubusercontent.com/29302909/68790995-756aa200-0659-11ea-8db5-f08520b2e3ea.png)

**8.** If users click **"start"** button the computation process is started.

![img5](https://user-images.githubusercontent.com/29302909/71669956-77123580-2d7f-11ea-959e-377bad1c6ada.png)

**9.** When the computation process is finished, a spreadsheet file which name is **"Synastry.xlsx"** is created.

## Spreadsheets

When all aspects and **"conjunction"** aspect is selected a table which is similar to below will be created. 

![img6](https://user-images.githubusercontent.com/29302909/68792282-1ce8d400-065c-11ea-9d1a-c61740c80930.jpeg)

When a planet like **"Sun"** and an aspect like **"conjunction"** are selected for detailed analysis, a table which is similar to below will be created.

![img7](https://user-images.githubusercontent.com/29302909/68792530-8ff24a80-065c-11ea-9b76-9b9bde8eb5ec.jpeg)

## Notes

**1.** The tables may not be opened by Microsoft Excel. Therefore it is recommended to use [Libre Office](https://www.libreoffice.org/download/download/). 

**2.** If users want to put their files in a cloud system like Dropbox, it is recommended that the format of the excel files should be changed from *xlsx* to *ods* format.

## Licenses

SynastryParser is released under the terms of the GNU GENERAL PUBLIC LICENSE. Please refer to the LICENSE file.
