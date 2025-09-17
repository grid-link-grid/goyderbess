\NeedsTeXFormat{LaTeX2e}
\ProvidesClass{grid-link-report}[2024/02/29 Grid-Link power systems report]

% Define an input parameter for the class
\DeclareOption*{\PassOptionsToClass{\CurrentOption}{book}}
\ProcessOptions\relax

% Define a command to set the assets directory
\newcommand{\SetClassAssetsDir}[1]{\renewcommand{\classassetsdir}{#1}}

% Set a default value for the assets directory
\newcommand{\classassetsdir}{../class-assets}

% Load the book class with the specified options
\LoadClass[12pt, oneside]{book}

\usepackage{pgfplotstable}
\RequirePackage{graphicx}
\RequirePackage[usenames,dvipsnames]{xcolor}
\RequirePackage{pdflscape}
\RequirePackage{ragged2e}
\RequirePackage{anyfontsize}
\RequirePackage{fancyhdr}
\RequirePackage{lipsum}
\RequirePackage{csvsimple-l3}
\RequirePackage{adjustbox}
\RequirePackage{titlesec}
\RequirePackage{marginnote}
\RequirePackage{booktabs}
\RequirePackage{longtable}
\RequirePackage{tabularx}
\RequirePackage{filecontents}
\RequirePackage{makecell}
\RequirePackage{underscore}
\RequirePackage{float}
\RequirePackage{amsmath}
\RequirePackage{hyperref}
\RequirePackage{fontspec}
\RequirePackage{tabularray}
\RequirePackage{colortbl}
\RequirePackage{sectsty}
\RequirePackage{lastpage}
\RequirePackage{parskip}
\RequirePackage{fixltx2e}
\RequirePackage{lscape}
\RequirePackage{appendix}
\RequirePackage{rotating}
\RequirePackage{parcolumns}
\RequirePackage{setspace}
\RequirePackage{xparse}
\RequirePackage{easy-todo}
\RequirePackage[printonlyused,withpage]{acronym}
\usepackage[skins,breakable]{tcolorbox}

\usepackage{lipsum}% example texts
\RequirePackage[stamp=false]{draftwatermark}
\RequirePackage[RPvoltages]{circuitikz}

\RequirePackage[
	a4paper, 
	margin=0.5in, 
	top=0.9in, 
	bottom=0.9in,
	footskip=0.4in,
]{geometry}



\usetikzlibrary{shapes,arrows,positioning,calc}

\hypersetup{
	colorlinks=false, %set true if you want colored links
	linktoc=all,     %set to all if you want both sections and subsections linked
	linkcolor=blue,  %choose some color if you want links to stand out
	pdfborder={0 0 0}, % Set link border to zero width
}


% Default paragraph skip size. Defined as a function so it can be re-called in environments where the skip size is reset, then called immediately.
\newcommand{\setparskip}{\setlength{\parskip}{1em}}
\setparskip

%\RequirePackage[margin=0in, paperwidth=8.27in, paperheight=11.69in]{geometry}
\definecolor{logogreen}{RGB}{100, 150, 75}
\definecolor{logoblue}{RGB}{49, 161, 191}
\definecolor{chapterheadingblue}{RGB}{39, 41, 54}
\definecolor{tableheaderblue}{RGB}{41, 49, 80}
\definecolor{tablesubheadergrey}{RGB}{127, 127, 127}
\definecolor{tablesubheaderblue}{RGB}{49, 161, 191}
\definecolor{tablefirstcolgrey}{RGB}{242, 242, 242}


% Variables
\newcommand{\project}[1]{\renewcommand{\project}{#1}}
\newcommand{\client}[1]{\renewcommand{\client}{#1}}
\newcommand{\developer}[1]{\renewcommand{\developer}{#1}}
\newcommand{\docnumber}[1]{\renewcommand{\docnumber}{#1}}
\newcommand{\issueddate}[1]{\renewcommand{\issueddate}{#1}}
\newcommand{\revision}[1]{\renewcommand{\revision}{#1}}
\newcommand{\revisionhistorycsvpath}[1]{\renewcommand{\revisionhistorycsvpath}{#1}}
\newcommand{\clientlogopath}[1]{\renewcommand{\clientlogopath}{#1}}


% Defined fonts
\setmainfont{Ubuntu}
\newcommand{\covertitlefont}{\fontsize{40}{42}\bfseries\selectfont\textcolor{logogreen}}
\newcommand{\coversubtitlefont}{\fontsize{20}{30}\selectfont\textcolor{black}}
\newcommand{\coverlabelfont}{\fontsize{15}{26}\selectfont\textcolor{black}}
\newcommand{\covervaluefont}{\fontsize{15}{26}\selectfont\textcolor{black}}
\newcommand{\footerfontbold}{\fontsize{13}{16}\bfseries\selectfont\textcolor{logogreen}}
\newcommand{\footerfontnormal}{\fontsize{13}{16}\selectfont\textcolor{logogreen}}


% Change how chapter titles display
\titleformat{\chapter}[hang] 
{\normalfont\huge\bfseries\color{chapterheadingblue}}{\thechapter.}{10pt}{} 
\titlespacing*{\chapter}{0pt}{0pt}{0pt} % Adjust as needed


\sectionfont{\color{logogreen}}  % sets colour of sections
\subsectionfont{\color{logogreen}}  % sets colour of sections

% Optional reversed acronym definition - e.g. "Virtual Machine (VM)" instead of "VM (Virtual Machine)"
\newcommand{\acreverse}[1]{\acsu{#1} (\aclu{#1})}

\newcommand*{\useportrait}{%
	\clearpage
	\KOMAoptions{paper=portrait,DIV=current}%switch to portrait
	\newgeometry{% geometry settings for portrait
		left=0.5in, right=0.5in, top=0.6in, bottom=1.25in,headheight=23pt,includehead
	}%
	\fancyhfoffset{0pt}% <- recalculate head and foot width for fancyhdr
}
\newcommand*{\uselandscape}{%
	\clearpage
	\KOMAoptions{paper=landscape,DIV=current}%switch to landscape
	\newgeometry{% geometry settings for landscape
		left=0.5in, right=0.5in, top=0.6in, bottom=1.2in,headheight=23pt,includehead, footskip=1in
	}%
	\fancyhfoffset{0pt}% recalculate head and foot width for fancyhdr
}

\fancypagestyle{plain}{%
	\fancyhf{}% clear all fields
	\renewcommand{\headrulewidth}{1.0pt} % Thickness of the line under the header
	\setlength{\headheight}{32pt}
	% Set the color of the headrule
	 \renewcommand{\headrule}{%
		\color{logogreen}\vskip0.1in% Add 1 inch of space between header and header rule
		\hrule width\headwidth height\headrulewidth \vskip-\headrulewidth%
	}%
	
	% Right side (logo)
	\fancyhead[R]{\includegraphics[width=1.2cm]{\classassetsdir/logo-simplified-small.png}}
	%\fancyhead[R]{\includegraphics[width=2.0cm]{\classassetsdir/logo-full-medium.png}}
	
	
	% Left side (two lines of left-aligned text)
	\fancyhead[L]{%
		\parbox[b]{\textwidth\color{logogreen}}{%
			\raggedright
			\project \\
			\@title
		}%
	}
	
	% Left side text in the footer
	\fancyfoot[L]{\color{logogreen}CONNECTING YOUR PROJECTS | GRID-LINK PTY LTD}
	% Right side page number in the footer for main matter pages
	\fancyfoot[R]{\if@mainmatter\thepage\ / \pageref{LastPage}\else\Roman{page}\fi}
}

% Apply the header style to all pages
\pagestyle{plain}

% Title page
\renewcommand{\maketitle}{%
	\newgeometry{margin=0in, paperwidth=8.27in, paperheight=11.69in}
	\thispagestyle{empty}
	\begin{titlepage}
		\vspace*{0.1in}
		
		\hspace*{0.05\linewidth}
		\begin{minipage}[c]{0.45\linewidth}
			\vspace*{0pt} % This is to align the top of the minipage with the top of the page
			\includegraphics[width=1.0in]{\clientlogopath}
		\end{minipage}%
		%\hspace*{0.05\linewidth}
		\begin{minipage}[c]{0.45\linewidth}
			\vspace*{0pt} % This is to align the top of the minipage with the top of the page
			\hfill % Add horizontal space
			\includegraphics[width=2.0in]{\classassetsdir/logo-full-medium.png}
		\end{minipage}
		\hspace*{0.05\linewidth}
		
		
		
		% Document number and client logo
		\hspace*{0.1\textwidth}
		
		
		\vspace*{0.7in}
		
		% Document title
		\begin{center}
			\covertitlefont{\project}
			
			\coversubtitlefont{\@title}
		\end{center}
		
		\vspace*{0.8in}
		
		% Document number and client logo
		\hspace*{0.1\textwidth}
		\begin{minipage}[t]{0.5\textwidth}
			\begin{raggedright}
				\begin{tabular}{@{}l@{\quad}l}
					\coverlabelfont{\textbf{Document:}} & \covervaluefont{\docnumber} \\
					\coverlabelfont{\textbf{Issue Date:}} & \covervaluefont{\issueddate} \\
					\coverlabelfont{\textbf{Revision:}} & \covervaluefont{\revision} \\
				\end{tabular}		
			\end{raggedright}
		\end{minipage}%
		\hfill
		
		\vfill
		
		% Bottom graphic
		\begin{minipage}{\textwidth}
			\centering
			\includegraphics[width=0.9\textwidth, keepaspectratio]{\classassetsdir/cover-lower-image.png}
		\end{minipage}
		
		\vspace*{0.2in}
		
		% Disclaimer
		\begin{minipage}{0.7\textwidth}
			\hspace*{0.3in}
			\footerfontbold{%
				CONNECTING YOUR PROJECTS |
			}
			\footerfontnormal{%
				grid-link.com.au
			}
			\vspace{0.5in}
		\end{minipage}
	\end{titlepage}
	\restoregeometry
}

% Title page
\newcommand{\makebackpage}{%
	\newgeometry{margin=0in, top=0in}
	\thispagestyle{empty}
	\includegraphics[width=\paperwidth]{\classassetsdir/cover-web-top.png}
	\vfill
	\begin{center}
		\includegraphics[width=1.6in]{\classassetsdir/logo-full-medium.png}
		
		\makeatletter
		\small{
			\textbf{Grid-Link Pty Ltd} | Level 12, 350 Collins St, Melbourne, VIC 3000 | ABN 55 651 392 746 | info@grid-link.com.au
		}
		\makeatother
	\end{center}
	\vspace{0.5in}
	\includegraphics[width=\paperwidth]{\classassetsdir/cover-web-bottom.png}
	\restoregeometry
}



\newcommand{\includepdfathreelandscape}[2]{%
	\begin{landscape}
		\newgeometry{left=-0.1in, right=0.5in, top=4in, bottom=1.2in}%
		\thispagestyle{empty}% Remove page number
		\foreach \page in {1,...,#1} {%
			\begin{figure}[p]

				\centering
				\includegraphics[page=\page,width=12in]{#2}
			\end{figure}
			\clearpage
		}
		\restoregeometry%
	\end{landscape}
}

% Title page
\newcommand{\makedisclaimer}{%
	\chapter{Disclaimer}
	Grid-Link disclaims responsibility to any person other than \client~arising in connection with this report. Grid-Link also excludes any warranties and conditions, to the extent legally permissible. The services undertaken by Grid-Link in connection with preparing this report were limited to those specifically detailed in the report and are subject to the scope limitations set out in the report. The opinions, conclusions and any recommendations in this report are based on conditions encountered and information reviewed at the date of preparation of the report. Grid-Link has no responsibility or obligation to update this report to account for events or changes occurring following the date that the report was prepared. The opinions, conclusions and any recommendations in this report are based on limitations and assumptions described in this report. Grid-Link disclaims liability arising from any of the assumptions being incorrect.
}

% About Grid-Link page
\newcommand{\makeaboutgridlink}{%
	\chapter{About Grid-Link}
	
	{%
		\setlength{\parskip}{1em} % Set paragraph spacing
		\setlength{\parindent}{0pt} % Disable paragraph indentation
		Grid-Link is a niche engineering firm providing end-to-end grid connection solutions. Solving the complexity of the grid is key to the energy transition. Grid-Link is passionate about reducing carbon emissions and enabling more clean energy projects in the electricity network.
		
		Grid-Link's expertise is unique. We combine deep technical understanding of grid modelling and AEMO compliance with practical experience in commissioning and testing as well as overall project development experience working for key EPC leaders in the industry.
		
		Our aim is to reduce complexity and risks for our clients.
		
		Our key differentiator is our expertise in successfully delivering complex energy projects working for network utilities, developers, contractors and market operator.
		Our services include:
		
		\begin{itemize}
			\item Grid connection
			\item Design
			\item End-to-end testing, commissioning and completions.
		\end{itemize}
	}
	
	

	
}



\newcommand{\autoscaledcolouredtable}[3]{%
	\begin{table}[#1]
		\centering
		\resizebox{\textwidth}{!}{%
			\begin{tabular}{|c|} \hline
				\rowcolor{gray!25} \bfseries Column Header \\ \hline
				\csvreader[
				late after line=\\\hline,
				before reading={\rowcolor{gray!10}\begingroup}, % Color rows
				after reading=\endgroup,
				head to column names
				]{#2}{}{\csvlinetotablerow}
			\end{tabular}
		}
	\end{table}%
}

\newcommand{\dansquicktable}[4]{	
	\thicktablelines
	\begin{longtable}{{#1}}%eg: {|C{5cm}|C{2cm}|C{2cm}|}
		\caption{#2}%eg: {FRT parameters}
		\label{tab:#3} \\% eg: {frt-parameters} \\
		\toprule
		\rowcolor{tableheaderblue}
		#4 \\%eg: {\bfseries \color{white}Parameter & \bfseries \color{white}Unit & \bfseries \color{white}Value} \\
		\endhead
		\bottomrule \endfoot
		\csvreader[
		late after line=\\\hline,
		late after last line=,
		before reading={\catcode`\#=12},
		after reading={\catcode`\#=6}]% 
		{#3}{}{\csvlinetotablerow} 
		\\\hline
	\end{longtable}
}

\newcommand{\dansquickertable}[2]
{\begin{table}
		\centering
		\thicktablelines
		%\caption{#2} % eg Dan's table
		%\label{tab:#3} % eg. dans-table (IT IS IMPORTANT that the csv is named the same way as the label for this command to work)
		\pgfplotstabletypeset[
		assign column name/.style={/pgfplots/table/column name={\textbf{\textcolor{white}{#1}}}}, %this has to stay as #1, so argument 1 must be fed in as #1
		col sep=comma,
		header=true,
		string type,
		% Style for the header row
		every head row/.style={
			before row={\toprule\hline\rowcolor{tableheaderblue}},
			after row=\hline,
		},
		% Style for body rows
		every odd row/.style={after row={\hline}},
		every even row/.style={after row={\hline}},
		every last row/.style={after row=\hline\bottomrule},
		% Column style to add cell borders
		every column/.style={
			column type={|c},},
		every last column/.style={column type={|c|}},
		format=file
		]{#2}% eg. report-assets\dans-table.csv
\end{table}}

%example of usage:
	%\dansquickertable{#1}%do not change, required for making the headings bold, white etc.
		%{Dan's table} %caption text
		%{dans-table} % csv file name, doubles as label for referencing

\newcommand{\dansreplacementauto}[4]
{\begin{table}
		\centering
		\thicktablelines
		\caption{#2} % eg Dan's table
		\label{tab:#3} % eg. dans-table 
		\pgfplotstabletypeset[
		assign column name/.style={/pgfplots/table/column name={\textbf{\textcolor{white}{#1}}}}, %this has to stay as #1, so argument 1 must be fed in as #1
		col sep=comma,
		header=true,
		string type,
		% Style for the header row
		every head row/.style={
			before row={\toprule\hline\rowcolor{tableheaderblue}},
			after row=\hline,
		},
		% Style for body rows
		every odd row/.style={after row={\hline}},
		every even row/.style={after row={\hline}},
		every last row/.style={after row=\hline\bottomrule},
		% Column style to add cell borders
		every column/.style={
			column type={|c},},
		every last column/.style={column type={|c|}},
		format=file
		]{#4}% eg. report-assets\dans-table.csv
\end{table}}

\newcommand{\coloredcell}[2][gray!10]{%
	\multicolumn{1}{|>{\columncolor{#1}}c|}{#2}%
}

% For very large tables
\newcommand{\autoscaledtable}[2]{%
	\begin{table}[#1]
		\centering
		\resizebox{\textwidth}{!}{%
			\csvautotabular{#2}
		}
	\end{table}%
}

% Table styles

% Increase the width of cell borders
\newcommand{\thicktablelines}{\setlength{\arrayrulewidth}{1.5pt}}

\newcommand{\autoscaledtablenew}[2]{%
	\csvstyle{revisiontablestyle}{
		tabular = |r|l|c|,
		table head = \hline & Person & Matr.~No.\\\hline\hline,
		late after line = \\\hline,
		head to column names,
		head to column names prefix = CSV,
	}
	
	\csvreader[revisiontablestyle]{grade.csv}{}{%
		\thecsvrow & \CSVgivenname~\CSVname & \CSVmatriculation
	}
}

% Left and right aligned wrapping column type
\newcolumntype{L}[1]{>{\raggedright\arraybackslash}p{#1}}
\newcolumntype{C}[1]{>{\centering\arraybackslash}p{#1}}
\newcolumntype{R}[1]{>{\raggedleft\arraybackslash}p{#1}}

\newcommand{\makerevisionhistorypage}{%
	\csvstyle{revisiontablestyle}{
		tabular = {|>{\columncolor{gray!20}}C{1cm}|C{2cm}|L{3.5cm}|L{3.5cm}|p{6cm}|},
		table head = \hline \rowcolor{tableheaderblue} \color{white}Rev. & \color{white}Date & \color{white}Prepared By & \color{white}Reviewed By & \color{white}Description \\\hline,
		late after line = \\\hline,
		head to column names,
		head to column names prefix = CSV,
	}

	\chapter{Revision History}
	\begin{table}[H]
		\centering
		\caption{Revision history}
		\label{tab:revisionhistory}
		\fontsize{11}{13}\selectfont
		\setlength{\arrayrulewidth}{1.5pt} % Adjust line width here
		\csvreader[revisiontablestyle]{\revisionhistorycsvpath}{}{%
			\CSVRevision & \CSVDate & \CSVPreparedBy & \CSVReviewedBy & \CSVDescription
		}
	\end{table}
	
	This document uses \emph{\href{https://semverdoc.org/semverdoc.html}{Semantic Versioning for Documents}} for revision numbering. 
	
	Given a version number \emph{MAJOR-MINOR-FIX}, the
	
	\begin{itemize}
		\item \emph{MAJOR} is incremented when the document has undergone significant changes
		\item \emph{MINOR} is incremented when new information has been added to the document or information has been removed from the document, and
		\item \emph{FIX} is incremented when minor changes are made (e.g. fixing typos)
	\end{itemize}
	
	Where appropriate, several revisions may be represented in one table entry with all notable changes described in the \emph{Description} column.
}

% Define a command to add or remove the draft stamp
\newcommand{\adddraftstamp}{
	\DraftwatermarkOptions{stamp=true}
	\SetWatermarkText{DRAFT}
	\SetWatermarkScale{1}
	\SetWatermarkColor[gray]{0.85}
}


% Notes
\newcommand{\discussionpointnum}{\bfseries\arabic{discussionpointcounter}}
\newcounter{discussionpointcounter}
\newcommand{\decrementdisccounter}{\addtocounter{discussionpointcounter}{-1}}

\newenvironment{disc}{%
	\stepcounter{discussionpointcounter}%
	\vspace*{20pt} 
	\begin{tcolorbox}[%
		blanker,%
		width=0.8\textwidth,%
		enlarge left by=0.15\textwidth,%
		before skip=6pt,%
		breakable,%
		overlay unbroken and first={%
			\node[inner sep=0pt,outer sep=0pt,text width=0.33\textwidth,align=none,below right]%
			at ([xshift=-0.12\textwidth]frame.north west)%
			{\textbf{Note \#\arabic{discussionpointcounter}}};}%
		]%
		\setlength{\parskip}{0.5em}
	}
	{%
	\end{tcolorbox}%
}



% Risks
\newcommand{\risknum}{\bfseries\arabic{riskcounter}}
\newcounter{riskcounter}
\newcommand{\decrementriskcounter}{\addtocounter{riskcounter}{-1}}


\newenvironment{risk}[1][]{%
	\ifthenelse{\equal{#1}{}}%
	{\stepcounter{riskcounter}\def\risknumber{\arabic{riskcounter}}}%
	{\def\risknumber{#1}}%
	\vspace*{20pt} 
	\colorlet{riskheadingorange}{Apricot}
	\begin{tcolorbox}[%
		blanker,%
		width=0.8\textwidth,%
		enlarge left by=0.15\textwidth,%
		before skip=6pt,%
		breakable,%
		overlay unbroken and first={%
			\node[inner sep=0pt,outer sep=0pt,text width=0.33\textwidth,align=none,below right]%
			at ([xshift=-0.12\textwidth]frame.north west)%
			{\textcolor{riskheadingorange}{\textbf{Risk \#\risknumber}}};}%
		]%
		\setlength{\parskip}{0.5em}
	}
	{%
	\end{tcolorbox}%
}

\newenvironment{highrisk}[1][]{%
	\ifthenelse{\equal{#1}{}}%
		{\stepcounter{riskcounter}\def\risknumber{\arabic{riskcounter}}}%
		{\def\risknumber{#1}}%
	\vspace*{20pt} 
	\colorlet{highriskheadingred}{red}
	\begin{tcolorbox}[%
		blanker,%
		width=0.8\textwidth,%
		enlarge left by=0.15\textwidth,%
		before skip=6pt,%
		breakable,%
		overlay unbroken and first={%
			\node[inner sep=0pt,outer sep=0pt,text width=0.33\textwidth,align=none,below right]%
			at ([xshift=-0.12\textwidth]frame.north west)%
			{\textcolor{highriskheadingred}{\textbf{Risk \#\risknumber}}};}%
		]%
		\setlength{\parskip}{0.5em}
	}
	{%
	\end{tcolorbox}%
}

\newenvironment{closedrisk}[1][]{%
	\ifthenelse{\equal{#1}{}}%
	{\stepcounter{riskcounter}\def\risknumber{\arabic{riskcounter}}}%
	{\def\risknumber{#1}}%
	\vspace*{20pt} 
	\colorlet{closedriskgrey}{gray}
	\begin{tcolorbox}[%
		blanker,%
		width=0.8\textwidth,%
		enlarge left by=0.15\textwidth,%
		before skip=6pt,%
		breakable,%
		overlay unbroken and first={%
			\node[inner sep=0pt,outer sep=0pt,text width=0.33\textwidth,align=none,below right]%
			at ([xshift=-0.12\textwidth]frame.north west)%
			{\textcolor{closedriskgrey}{\textbf{Risk \#\risknumber}}};}%
		]%
		\setlength{\parskip}{0.5em}
	}
	{%
	\end{tcolorbox}%
}



% Assumptions (Same as dicussion points)
\newcommand{\assumptionnum}{\bfseries\arabic{assumptioncounter}}
\newcounter{assumptioncounter}
\newcommand\assumption{%
	\stepcounter{assumptioncounter}%
	\vspace*{10pt} % Add 40pt vertical space
	\noindent
	% Optional: Use the below line to put #n in the margins
	%\leavevmode\marginpar[\hfill\assumptionnum]{\#\parnum}%
	\textbf{(\arabic{assumptioncounter})}\hspace{10pt}%
}

% Risk matrix
\newcommand{\makeriskmatrix}{%
	\thicktablelines
	\colorlet{lowrisk}{LimeGreen}
	\colorlet{mediumrisk}{yellow}
	\colorlet{highrisk}{Apricot}
	\colorlet{extremerisk}{red}
	\begin{tabular}{|c|c|c|c|c|c|c|}
		\multicolumn{1}{c|}{} & \multicolumn{5}{|c|}{\cellcolor{tableheaderblue}\textcolor{white}{\textbf{Consequence}}} \\
		\hline
		\cellcolor{tableheaderblue}\textcolor{white}{\textbf{Likelihood}} & \cellcolor{tablesubheaderblue}\textcolor{white}{\textbf{Negligible}} & \cellcolor{tablesubheaderblue}\textcolor{white}{\textbf{Minor}} & \cellcolor{tablesubheaderblue}\textcolor{white}{\textbf{Moderate}} & \cellcolor{tablesubheaderblue}\textcolor{white}{\textbf{Major}} & \cellcolor{tablesubheaderblue}\textcolor{white}{\textbf{Severe}} \\
		\hline
		\cellcolor{tablesubheaderblue}\textcolor{white}{\textbf{Almost Certain}} & \cellcolor{mediumrisk} Medium & \cellcolor{highrisk} High & \cellcolor{highrisk} High & \cellcolor{extremerisk} Extreme & \cellcolor{extremerisk} Extreme \\
		\hline
		\cellcolor{tablesubheaderblue}\textcolor{white}{\textbf{Likely}} & \cellcolor{mediumrisk} Medium & \cellcolor{mediumrisk} Medium & \cellcolor{highrisk} High & \cellcolor{highrisk} High & \cellcolor{extremerisk} Extreme \\
		\hline
		\cellcolor{tablesubheaderblue}\textcolor{white}{\textbf{Probable}} & \cellcolor{lowrisk} Low & \cellcolor{mediumrisk} Medium & \cellcolor{mediumrisk} Medium & \cellcolor{highrisk} High & \cellcolor{extremerisk} Extreme \\
		\hline
		\cellcolor{tablesubheaderblue}\textcolor{white}{\textbf{Unlikely}} & \cellcolor{lowrisk} Low & \cellcolor{mediumrisk} Medium & \cellcolor{mediumrisk} Medium & \cellcolor{mediumrisk} Medium & \cellcolor{highrisk} High \\
		\hline
		\cellcolor{tablesubheaderblue}\textcolor{white}{\textbf{Rare}} & \cellcolor{lowrisk} Low & \cellcolor{lowrisk} Low & \cellcolor{mediumrisk} Medium & \cellcolor{mediumrisk} Medium & \cellcolor{highrisk} High \\
		\hline
	\end{tabular}
}

% Risk likelihood definitions
\newcommand{\makelikelihooddefinitionstable}{%
	\definecolor{likelihoodgrey}{gray}{0.85}
	\thicktablelines
	\begin{tabular}{|c|c|c|}
		\hline
		\rowcolor{tableheaderblue} \textcolor{white}{\textbf{Likelihood}} & \textcolor{white}{\textbf{Rating}} & \textcolor{white}{\textbf{Meaning}} \\
		\hline
		\cellcolor{likelihoodgrey} Almost Certain & \cellcolor{likelihoodgrey} 5 & Expected in most circumstances (>90\%) \\
		\hline
		\cellcolor{likelihoodgrey} Likely & \cellcolor{likelihoodgrey} 4 & Will probably occur in most circumstances (75\%) \\
		\hline
		\cellcolor{likelihoodgrey} Possible & \cellcolor{likelihoodgrey} 3 & Could occur at some time (50\%) \\
		\hline
		\cellcolor{likelihoodgrey} Unlikely & \cellcolor{likelihoodgrey} 2 & Not expected to occur (25\%) \\
		\hline
		\cellcolor{likelihoodgrey} Rare & \cellcolor{likelihoodgrey} 1 & Exceptional circumstances only (<10\%) \\
		\hline
	\end{tabular}
}

% Risk likelihood definitions
\newcommand{\makeconsequencedefinitionstable}{%
	\definecolor{consequencegrey}{gray}{0.85}
	\thicktablelines
	\begin{tabular}{|c|c|c|}
		\hline
		\rowcolor{tableheaderblue} \textcolor{white}{\textbf{Consequence}} & \textcolor{white}{\textbf{Rating}} & \textcolor{white}{\textbf{Meaning}} \\
		\hline
		\cellcolor{consequencegrey} Severe & \cellcolor{consequencegrey} 5 & \$1M+ cost, more than 6 month delay or indefinite stalemate \\
		\hline
		\cellcolor{consequencegrey} Major & \cellcolor{consequencegrey} 4 & \$500k-1M cost or 6 month delay \\
		\hline
		\cellcolor{consequencegrey} Moderate & \cellcolor{consequencegrey} 3 &\$200-500k cost or 3 month delay \\
		\hline
		\cellcolor{consequencegrey} Minor & \cellcolor{consequencegrey} 2 & \$50-200k cost or 1 month delay \\
		\hline
		\cellcolor{consequencegrey} Negligible & \cellcolor{consequencegrey} 1 & <\$50k cost or <1 month delay \\
		\hline
	\end{tabular}
}

% Risk likelihood definitions
\newcommand{\makeriskmeaningstable}{%
	\colorlet{negligiblerisk}{SkyBlue}
	\colorlet{lowrisk}{LimeGreen}
	\colorlet{mediumrisk}{yellow}
	\colorlet{highrisk}{Apricot}
	\colorlet{extremerisk}{red}
	\thicktablelines
	\begin{tabular}{|c|C{12cm}|}
		\hline
		\rowcolor{tableheaderblue} \textcolor{white}{\textbf{Risk Level}} & \textcolor{white}{\textbf{Risk Action}} \\
		\hline
		\cellcolor{extremerisk} Extreme & Report risk and status to the working group at scheduled meetings. Requires the immediate attention of Client. Undertake detailed research, including the identification of additional risk reduction options and preparation of a detailed risk management plan. Regular monitoring of risk reduction strategies. \\
		\hline
		\cellcolor{highrisk} High & Report risk and status to the working group at scheduled meetings. Undertake detailed research, including the identification of additional risk reduction options and preparation of a detailed risk management plan. Regular monitoring of risk reduction strategies. \\
		\hline
		\cellcolor{mediumrisk} Medium & Responsibility for management resides with the relevant workstream leader and specific monitoring of response procedures should occur \\
		\hline
		\cellcolor{lowrisk} Low & Look to accept, monitor and report on the risk or manage it through the normal management practices \\
		\hline
	\end{tabular}
}

\endinput
