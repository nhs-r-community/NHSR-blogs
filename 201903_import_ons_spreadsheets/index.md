---
title: "Format ONS spreadsheet"
output: hugodown::md_document
rmd_hash: 75cbf9da8751fbbf

---

# Background

A Public Health consultant colleague Ian Bowns (@IantheBee) created a report to monitor mortality within the Trust and he used the ONS weekly provisional data for the East Midlands to compare the pattern and trends of deaths over time. This involves downloading a file from:

<https://www.ons.gov.uk/peoplepopulationandcommunity/birthsdeathsandmarriages/deaths/datasets/weeklyprovisionalfiguresondeathsregisteredinenglandandwales>

which is updated weekly. Once a month I, manually, add numbers from this to another spreadsheet to be imported to R for the overall analysis.

# Downloaded file formats

You may be familiar with ONS and other NHS data spreadsheets format and if you are not, here are some of the issues:

-   data is presented in wide form, not long form (so columns are rows and vice versa)
-   the sheets are formatted for look rather than function with numerous blank rows and blank columns
-   there are multiple sheets with information about methodology usually coming first. This means a default upload to programmes like R are not possible as they pick up the first sheet/tab
-   the file name changes each week and includes the date which means any code to pick up a file needs to be changed accordingly for each load
-   being Excel, when this is imported into R, there can be problems with the date formats. These can get lost to the Excel Serial number and
-   they include a lot of information and often only need a fraction of it

Given these difficulties there is great temptation, as happened with this, to just copy and paste what you need. This isn't ideal for the reasons:

-   it increases the likelihood of data entry input error
-   it takes time and
-   it is just so very very tedious

The solution is, always, to automate and tools like Power Pivot in Excel or SSIS could work but as the final report is in R it makes sense to tackle this formatting in R and this is the result...

# Import file

For this you can either save the file manually or use the following code within R. Save it to the same place where the code is running and you should see the files in the bottom right window under the tab 'Files'. The best way to do this is using project and opening up the script within that project.

<div class="highlight">

<pre class='chroma'><code class='language-r' data-lang='r'><span></span>
<span><span class='nf'><a href='https://rdrr.io/r/utils/download.file.html'>download.file</a></span><span class='o'>(</span><span class='s'>"https://www.ons.gov.uk/file?uri=/peoplepopulationandcommunity/birthsdeathsandmarriages/deaths/datasets/weeklyprovisionalfiguresondeathsregisteredinenglandandwales/2019/publishedweek082019.xls"</span>,</span>
<span>              destfile <span class='o'>=</span> <span class='s'>"DeathsDownload.xls"</span>,</span>
<span>              method <span class='o'>=</span> <span class='s'>"wininet"</span>, <span class='c'>#use "curl" for OS X / Linux, "wininet" for Windows</span></span>
<span>              mode <span class='o'>=</span> <span class='s'>"wb"</span><span class='o'>)</span> <span class='c'>#wb means "write binary"</span></span>
<span><span class='c'>#&gt; Warning in download.file("https://www.ons.gov.uk/file?uri=/peoplepopulationandcommunity/birthsdeathsandmarriages/deaths/datasets/weeklyprovisionalfiguresondeathsregisteredinenglandandwales/2019/publishedweek082019.xls", : the 'wininet' method is deprecated for http:// and https:// URLs</span></span>
<span></span></code></pre>

</div>

Not that this file's name and URL changes *each* week so the code needs changing *each* time it is run.

Once the file is saved use readxl to import which means the file doesn't need its format changing from the original .xls

When I upload this file I get warnings which are related, I think, to the Excel serial numbers appearing where dates are expected.

-   sheet = :refers to the sheet I want. I think this has to be numeric and doesn't use the tab's title.
-   skip = :is the number of top rows to ignore.

<div class="highlight">

<pre class='chroma'><code class='language-r' data-lang='r'><span></span>
<span><span class='nv'>DeathsImport</span> <span class='o'>&lt;-</span> <span class='nf'><a href='https://readxl.tidyverse.org/reference/read_excel.html'>read_excel</a></span><span class='o'>(</span><span class='s'>"DeathsDownload.xls "</span>, </span>
<span>                           sheet <span class='o'>=</span> <span class='m'>4</span>,</span>
<span>                           skip <span class='o'>=</span> <span class='m'>2</span><span class='o'>)</span></span>
<span><span class='c'>#&gt; Warning: Expecting numeric in C5 / R5C3: got a date</span></span>
<span></span><span><span class='c'>#&gt; Warning: Expecting numeric in F5 / R5C6: got a date</span></span>
<span></span><span><span class='c'>#&gt; Warning: Expecting numeric in G5 / R5C7: got a date</span></span>
<span></span><span><span class='c'>#&gt; Warning: Expecting numeric in H5 / R5C8: got a date</span></span>
<span></span><span><span class='c'>#&gt; Warning: Expecting numeric in I5 / R5C9: got a date</span></span>
<span></span><span><span class='c'>#&gt; New names:</span></span>
<span><span class='c'>#&gt; <span style='color: #00BBBB;'>•</span> `` -&gt; `...2`</span></span>
<span></span></code></pre>

</div>

# Formatting the data

The next code creates a list that is used in the later code that is similar to the SQL IN but without typing out the list *within* the code for example:

-   SQL : WHERE city IN ('Paris','London','Hull')
-   R : filter(week_number %in% filter)

These lines of code are base R code and so don't rely on any packages.

<div class="highlight">

<pre class='chroma'><code class='language-r' data-lang='r'><span></span>
<span><span class='nv'>LookupList</span> <span class='o'>&lt;-</span> <span class='nf'><a href='https://rdrr.io/r/base/c.html'>c</a></span><span class='o'>(</span><span class='s'>"Week ended"</span>,</span>
<span>            <span class='s'>"Total deaths, all ages"</span>,</span>
<span>            <span class='s'>"Total deaths: average of corresponding"</span>,</span>
<span>            <span class='s'>"E12000004"</span></span>
<span><span class='o'>)</span></span></code></pre>

</div>

The next bit uses the dplyr package, which has loaded as part of tidyverse, as well as the janitor package. Not all packages are compatible with tidyverse but many do as this is often the go-to data manipulation package.

As an aside the %\>% is called a pipe and the shortcut is Shift + Ctrl + m. Worth learning as you'll be typing a lot more if you type out those pipes each time.

Janitor commands

-   Clean names: removes spaces in column headers and replaces with \_
-   remove_empty: gets rid of rows and columns - this dataset has a lot of those!

Dplyr command

-   filter: is looking just for the rows with the words from the list 'LookupList'. These will become the column names later.

<div class="highlight">

<pre class='chroma'><code class='language-r' data-lang='r'><span></span>
<span><span class='nv'>DeathsImport2</span> <span class='o'>&lt;-</span> <span class='nv'>DeathsImport</span> <span class='o'><a href='https://magrittr.tidyverse.org/reference/pipe.html'>%&gt;%</a></span> </span>
<span>  <span class='nv'>clean_names</span> <span class='o'><a href='https://magrittr.tidyverse.org/reference/pipe.html'>%&gt;%</a></span> </span>
<span>  <span class='nf'><a href='https://sfirke.github.io/janitor/reference/remove_empty.html'>remove_empty</a></span><span class='o'>(</span><span class='nf'><a href='https://rdrr.io/r/base/c.html'>c</a></span><span class='o'>(</span><span class='s'>"rows"</span>,<span class='s'>"cols"</span><span class='o'>)</span><span class='o'>)</span> <span class='o'><a href='https://magrittr.tidyverse.org/reference/pipe.html'>%&gt;%</a></span> </span>
<span>  <span class='nf'><a href='https://dplyr.tidyverse.org/reference/filter.html'>filter</a></span><span class='o'>(</span><span class='nv'>week_number</span> <span class='o'><a href='https://rdrr.io/r/base/match.html'>%in%</a></span> <span class='nv'>LookupList</span><span class='o'>)</span> </span></code></pre>

</div>

There are great commands called gather and spread which can be used to move wide form data to long and vice versa but with this I noticed that I just needed to turn it on its side so I used t() which is also useful as it turns the data frame to a matrix. You can see this by looking in the 'Environment' window in the top right of R Studio; there is no blue circle with an arrow next to t_DeathsImport.

<div class="highlight">

<pre class='chroma'><code class='language-r' data-lang='r'><span></span>
<span><span class='nv'>t_DeathsImport</span> <span class='o'>&lt;-</span> <span class='nf'><a href='https://rdrr.io/r/base/t.html'>t</a></span><span class='o'>(</span><span class='nv'>DeathsImport2</span><span class='o'>)</span></span></code></pre>

</div>

Being a matrix is useful as the next line of code makes the first row into column headers and this only works on a matrix.

<div class="highlight">

<pre class='chroma'><code class='language-r' data-lang='r'><span></span>
<span><span class='nf'><a href='https://rdrr.io/r/base/colnames.html'>colnames</a></span><span class='o'>(</span><span class='nv'>t_DeathsImport</span><span class='o'>)</span> <span class='o'>&lt;-</span> <span class='nv'>t_DeathsImport</span><span class='o'>[</span><span class='m'>1</span>, <span class='o'>]</span></span></code></pre>

</div>

Dplyr gives an error on matrices:

Code:

t_DeathsImport %\>% mutate(serialdate = excel_numeric_to_date(as.numeric(as.character(`Week ended`)), date_system = "modern"))

Result:

Error in UseMethod("mutate\_") : no applicable method for 'mutate\_' applied to an object of class "c('matrix', 'character')"

As later code will need dplyr turn the matrix into a dataframe using some base R code:

<div class="highlight">

<pre class='chroma'><code class='language-r' data-lang='r'><span></span>
<span><span class='nv'>t_DeathsImport</span> <span class='o'>&lt;-</span> <span class='nf'><a href='https://rdrr.io/r/base/as.data.frame.html'>as.data.frame</a></span><span class='o'>(</span><span class='nv'>t_DeathsImport</span><span class='o'>)</span></span></code></pre>

</div>

Previous dplyr code filtered on an %in% bit of code and it's natural to want a %not in% but it doesn't exist! However, cleverer minds have worked out a function:

<div class="highlight">

<pre class='chroma'><code class='language-r' data-lang='r'><span></span>
<span><span class='s'>'%!ni%'</span> <span class='o'>&lt;-</span> <span class='kr'>function</span><span class='o'>(</span><span class='nv'>x</span>,<span class='nv'>y</span><span class='o'>)</span><span class='o'>!</span><span class='o'>(</span><span class='s'>'%in%'</span><span class='o'>(</span><span class='nv'>x</span>,<span class='nv'>y</span><span class='o'>)</span><span class='o'>)</span></span></code></pre>

</div>

The text between the '' can be anything but I like '%ni%' as it's reminiscent of Monty Python.

Because of the moving around of rows to columns the data frame now has a row of column names which is not necessary as well as a row with just 'East Midlands' in one of the columns so the following 'remove' list is a workaround to get rid of these two lines.

<div class="highlight">

<pre class='chroma'><code class='language-r' data-lang='r'><span></span>
<span><span class='nv'>remove</span> <span class='o'>&lt;-</span> <span class='nf'><a href='https://rdrr.io/r/base/c.html'>c</a></span><span class='o'>(</span><span class='s'>"E12000004"</span>, <span class='s'>"East Midlands"</span><span class='o'>)</span></span></code></pre>

</div>

The next code uses the above list followed by a mutate which is followed by a janitor command 'excel_numeric_to_date'. This tells it like it is but, as often happens, the data needs to be changed to a character and *then* to numeric. The date system = "modern" isn't needed for this data but as I took this from the internet and it worked, so I left it.

An error will appear about NAs (nulls).

<div class="highlight">

<pre class='chroma'><code class='language-r' data-lang='r'><span></span>
<span><span class='nv'>t_DeathsImport</span> <span class='o'>&lt;-</span> <span class='nv'>t_DeathsImport</span> <span class='o'><a href='https://magrittr.tidyverse.org/reference/pipe.html'>%&gt;%</a></span> </span>
<span>  <span class='nf'><a href='https://dplyr.tidyverse.org/reference/filter.html'>filter</a></span><span class='o'>(</span><span class='nv'>E12000004</span> <span class='o'>%!ni%</span> <span class='nv'>remove</span><span class='o'>)</span> <span class='o'><a href='https://magrittr.tidyverse.org/reference/pipe.html'>%&gt;%</a></span> </span>
<span>  <span class='nf'><a href='https://dplyr.tidyverse.org/reference/mutate.html'>mutate</a></span><span class='o'>(</span>serialdate <span class='o'>=</span> <span class='nf'><a href='https://sfirke.github.io/janitor/reference/excel_numeric_to_date.html'>excel_numeric_to_date</a></span><span class='o'>(</span><span class='nf'><a href='https://rdrr.io/r/base/numeric.html'>as.numeric</a></span><span class='o'>(</span><span class='nf'><a href='https://rdrr.io/r/base/character.html'>as.character</a></span><span class='o'>(</span><span class='nv'>`Week ended`</span><span class='o'>)</span><span class='o'>)</span>, date_system <span class='o'>=</span> <span class='s'>"modern"</span><span class='o'>)</span><span class='o'>)</span></span>
<span><span class='c'>#&gt; Warning: There was 1 warning in `mutate()`.</span></span>
<span><span class='c'>#&gt; <span style='color: #00BBBB;'>ℹ</span> In argument: `serialdate = excel_numeric_to_date(...)`.</span></span>
<span><span class='c'>#&gt; Caused by warning in `excel_numeric_to_date()`:</span></span>
<span><span class='c'>#&gt; <span style='color: #BBBB00;'>!</span> NAs introduced by coercion</span></span>
<span></span></code></pre>

</div>

Now to deal with this mixing of real dates with Excel serial numbers.

Firstly, the following code uses base R to confirm real dates are real dates which conveniently wipes the serial numbers and makes them NAs.

<div class="highlight">

<pre class='chroma'><code class='language-r' data-lang='r'><span></span>
<span><span class='nv'>t_DeathsImport</span><span class='o'>$</span><span class='nv'>`Week ended`</span> <span class='o'>&lt;-</span> <span class='nf'><a href='https://rdrr.io/r/base/as.Date.html'>as.Date</a></span><span class='o'>(</span><span class='nv'>t_DeathsImport</span><span class='o'>$</span><span class='nv'>`Week ended`</span>, format <span class='o'>=</span> <span class='s'>'%Y-%m-%d'</span><span class='o'>)</span></span></code></pre>

</div>

This results in two columns:

-   `Week ended` which starts off with NAs then becomes real dates and
-   serialdate which starts off with real dates and then NAs.

The human eye and brain can see that these two follow on from each other and just, somehow, need to be squished together and the code to do it is as follows:

<div class="highlight">

<pre class='chroma'><code class='language-r' data-lang='r'><span></span>
<span><span class='nv'>t_DeathsImport</span> <span class='o'>&lt;-</span> <span class='nv'>t_DeathsImport</span> <span class='o'><a href='https://magrittr.tidyverse.org/reference/pipe.html'>%&gt;%</a></span> </span>
<span>  <span class='nf'><a href='https://dplyr.tidyverse.org/reference/mutate.html'>mutate</a></span><span class='o'>(</span>date <span class='o'>=</span> <span class='nf'><a href='https://dplyr.tidyverse.org/reference/if_else.html'>if_else</a></span><span class='o'>(</span><span class='nf'><a href='https://rdrr.io/r/base/NA.html'>is.na</a></span><span class='o'>(</span><span class='nv'>`Week ended`</span><span class='o'>)</span>,<span class='nv'>serialdate</span>,<span class='nv'>`Week ended`</span><span class='o'>)</span><span class='o'>)</span></span></code></pre>

</div>

To translate the mutate, this creates a new column called date which, if the `Week ended` is null then takes the serial date, otherwise it takes the `Week ended`.

Interestingly if 'ifelse' without the underscore is used it converts the dates to integers and these are not the same as the Excel serial numbers so use 'if_else'!

And that's it.

Or is it?

You might want to spit out the data frame *back* into excel and that's where a different package called openxlsx can help. As with many things with R, "other packages are available".

<div class="highlight">

<pre class='chroma'><code class='language-r' data-lang='r'><span></span>
<span><span class='nf'><a href='https://rdrr.io/pkg/openxlsx/man/write.xlsx.html'>write.xlsx</a></span><span class='o'>(</span><span class='nv'>DeathsImport</span>, <span class='s'>'ImportProvisionalDeaths.xlsx'</span><span class='o'>)</span></span></code></pre>

</div>

If you haven't used a project (which is really the best way to work) this will probably save in some obscure C: drive that you'll see in the bottom left 'Console' just under the tab names for 'Console' and 'Terminal'. Using projects means you set the pathway and that will mean the file saves in the same place and will also appear in the bottom *right* panel under 'Files'.

# Feedback

I'm pretty early on in my journey in R and many of my colleagues still haven't started yet so I'm throwing this out there so everyone can see it, newbies and old hands alike. If you spot anything, can explain anything further, need more explanation or can offer any alternatives to what I've done please please feel free to comment.

