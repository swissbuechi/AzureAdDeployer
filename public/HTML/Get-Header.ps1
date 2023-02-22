function Get-Header {
    return @"
<title>$($script:ReportTitle)</title>
<link rel="icon" type="image/png" href="$($script:ReportImageUrl)">
<style>
html {
    display: table;
    margin: auto;
}
body {
    display: table-cell;
    vertical-align: middle;
    padding-right: 200px;
    padding-left: 200px;
}
h1 {
    font-family: Arial, Helvetica, sans-serif;
    color: #666666;
    font-size: 32px;
}
h2 {
    font-family: Arial, Helvetica, sans-serif;
    color: #666666;
    font-size: 24px;
}
h3 {
    font-family: Arial, Helvetica, sans-serif;
    color: #666666;
    font-size: 16px;

}
p {
    font-family: Arial, Helvetica, sans-serif;
    font-size: 14px;
}
a {
    font-family: Arial, Helvetica, sans-serif;
    font-size: 16px;
    text-decoration: none;
    color: #666666;
}
ul {
    list-style-type: none;
    margin-top: 5px;
}
li {
    padding: 5px;
}
table {
    font-size: 14px;
    border: 0px;
    font-family: Arial, Helvetica, sans-serif;
    border-collapse: collapse;
    margin: 25px 0;
    min-width: 400px;
    box-shadow: 0 0 20px rgba(0, 0, 0, 0.15);
}
th,
td {
    padding: 4px;
    margin: 0px;
    border: 0;
    padding: 12px 15px;
}
th {
    background: #666666;
    color: #fff;
    font-size: 11px;
    padding: 10px 15px;
    vertical-align: middle;
}
tbody tr:nth-child(even) {
    background: #f0f0f2;
}
thead tr {
    color: #ffffff;
    text-align: left;
}
tbody tr {
    border-bottom: 1px solid #dddddd;
}
tbody tr:nth-of-type(even) {
    background-color: #f3f3f3;
}
.red {
    color: red;
}
.orange {
    color: orange;
}
.TOC {
    margin: 5px;
}
#FootNote {
font-family: Arial, Helvetica, sans-serif;
color: #666666;
font-size: 12px;
}
</style>
"@
}