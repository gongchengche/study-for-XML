package require opencv
package require Img
package require BWidget
package require mupdf
package require mupdf::widget
package require math::geometry
package require cawt
package require tcom

# tool GUI
variable width
variable height
variable img
variable w
variable v ""
variable outVideo
variable frameNum 0
variable currentPage 0
variable videoPath ""

proc buildGUI {} {
	global mainframe pages PDFhandle planPDFhandle commentsarear videoPath
	set descmenu {
		"&File" {} {} 0 {
	        {command "&New"     {} "Create a new Test Project"     {Ctrl n} -command Menu::new}
	        {command "&Open..." {} "Open an existing Project" {Ctrl o} -command Menu::open}
	        {command "&Save"    open "Save the Project" {Ctrl s} -command Menu::save}
	        {separator}
	        {cascade "&Recent files" {} recent 0 {}}
	        {separator}
	        {command "E&xit" {} "Exit the application" {} -command Menu::exit}
	    }
	    "&Configure" {} {} 0 {
	    	{command "Load configuration excel" {} "Load configuration excel"     {} -command loadConfiguration}
	    }
	    "&Test" {} {} 0 {
	    	{command "Open PDF" {} "Open PDF"     {} -command openPDF}
	    	{command "Close PDF" {} "Close PDF"     {} -command closePDF}
	    	{command "Load plan" {} "Load plan"     {} -command openPlanPDF}
	    	{command "Close plan" {} "Close plan"     {} -command closePlanPDF}
	    	{command "Record video" {} "Record video"     {} -command videoCapture}
	    	{command "Stop Record" {} "Stop record video"     {} -command videoCaptureStop}
	    }
	    "&Replay" {} {} 0 {
	    	{command "Play video" {} "Play video"     {} -command videoPlay}
	    	{command "Pause video" {} "Pause video"     {} -command videoPause}
	    	{command "Stop video" {} "Stop video"     {} -command videoStop}
	    }
	    "&Repot" {} {} 0 {
	    	{command "Generate repot" {} "Generate repot"     {} -command fillReport}
	    }
	   	"&Help" {} {} 0 {
	    }
	}

	MainFrame .main -menu $descmenu
	set mainframe [.main getframe]
	set pwh [panedwindow $mainframe.pwh -opaqueresize false]
	set pwv1 [panedwindow $mainframe.pwv1 -orient vertical -opaqueresize false]
	set pwv2 [panedwindow $mainframe.pwv2 -orient vertical -opaqueresize false]
	$pwh add $pwv1
	$pwh add $pwv2
	
	canvas $mainframe.cplan -bg darkgrey
	image create photo videoImg
    label $mainframe.lb -image videoImg -bg darkgrey

	frame $mainframe.testframe -bg darkgrey
	frame $mainframe.comments -bg darkgrey
 
	$pwv1 add $mainframe.lb 
	$pwv1 add $mainframe.cplan 
	$pwv2 add $mainframe.testframe
	$pwv2 add $mainframe.comments

	pack $pwh -fill both -expand yes
	pack .main -fill both -expand yes
	
	# 创建文本框右键菜单
	set commentsMenu [menu $mainframe.comments.commentsMenu -tearoff 0]
	$commentsMenu add command -label "Record comments" -command commentSave
	$commentsMenu add command -label "Alter comments" -command alterComments

	#创建评论区
	set commentsarear [listbox $mainframe.comments.commentsbox -bg darkgrey]
	bind $commentsarear <Button-3> {set x [expr %X]; set y [expr %Y]; tk_popup $mainframe.comments.commentsMenu $x $y}
	
	#创建滑动条
	set sy [scrollbar $mainframe.comments.sy -orient vertical -command [list $commentsarear yview]]
	set sx [scrollbar $mainframe.comments.sx -orient horizontal -command [list $commentsarear xview]]
	$commentsarear configure -yscrollcommand [list $sy set] -xscrollcommand [list $sx set]
	grid $commentsarear $sy -sticky news
	grid $sx            -   -sticky news
	grid rowconfigure $mainframe.comments 0 -weight 1
	grid columnconfigure $mainframe.comments 0 -weight 1
	updatecomments $commentsarear

	# some window setting (fixme, better solution to find):
	wm geometry . 1920x991+-14+0
	update
	$mainframe.pwh sash place 0 1090 1
	$mainframe.pwv1 sash place 0 1 665
	$mainframe.pwv2 sash place 0 1 780

	global textSearch {}
	
	# page up/down
	bind MuPdfWidget <Key-Next>  {
		if {$currentPage < $pages} {
			%W nextpage
			incr currentPage
		}
	}
	bind MuPdfWidget <Key-Prior> {
		if {$currentPage ne "0"} {
			%W prevpage
			incr currentPage -1
		}
	}
    bind MuPdfWidget <MouseWheel> { %W scroll 0 [expr -%D] }

	# ... it's enough to adjust the x/y-scrollincrement  .. be careful: integer, >=1 ...
	bind MuPdfWidget <Key-Up>    { %W scroll 0 -10 }
	bind MuPdfWidget <Key-Down>  { %W scroll 0 +10 }
	bind MuPdfWidget <Key-Left>  { %W scroll -10 0 }				
	bind MuPdfWidget <Key-Right> { %W scroll +10 0 }
	
	# Button-1 and drag, scroll the page
	bind MuPdfWidget <ButtonPress-1> { focus %W ; %W scan mark %x %y }
	bind MuPdfWidget <B1-Motion> { %W scan dragto %x %y 1 }

	# zoom +/-
	bind MuPdfWidget <Key-plus>  { %W rzoom +1 }
	bind MuPdfWidget <Key-minus> { %W rzoom -1 }
	bind MuPdfWidget <Key-x> { %W zoomfit x }
	bind MuPdfWidget <Key-y> { %W zoomfit y }
	bind MuPdfWidget <Key-z> { %W zoomfit xy }
	
	bind MuPdfWidget <Control-Key-f> { openSearchPanel %W }

	# 绑定窗口关闭事件
	bind $mainframe <Destroy> {
	    # 保存内容到文件
	    if {[info exist PDFhandle]} {
			catch {$PDFhandle close}
		}

		if {[info exist planPDFhandle]} {
			catch {$planPDFhandle close}
		}

	    # 退出程序
	    exit
	}
}

# 加载配置文件
proc loadConfiguration {} {
	set confExcelPath [tk_getOpenFile]
	try {
		set app [Excel OpenNew false]
		catch {Excel ShowAlerts $app off}
		set wb [Excel OpenWorkbook $app $confExcelPath]
		set wsNum [Excel GetNumWorksheets $wb]
		set ws [Excel GetWorksheetIdByName $wb "PDF Configure"]
		set matrix [Excel GetWorksheetAsMatrix $ws]
		foreach line $matrix {
			if {[lindex $line 0]=="Test Procedure"} {
				set testProcedurePath [file join [pwd] input "[lindex $line 1].pdf"]
				if {![file exists $testProcedurePath]} {puts "file $testProcedurePath not exist"; continue}
				openPDF $testProcedurePath
			} elseif {[lindex $line 0]=="Scheme Plan"} {
				set schemePlanPath [file join [pwd] input "[lindex $line 1].pdf"]
				if {![file exists $schemePlanPath]} {puts "file $schemePlanPath not exist"; continue}
				loadPicture2excel $schemePlanPath
			}
		}
	} on error {em} {
		puts $em
	}
	Excel Close $wb
	Cawt Destroy $app
}

# 加载图片到excel
proc loadPicture2excel {{picturefile ""}} {
	if {$picturefile== ""} {set picturefile [tk_getOpenFile] }
	if {$picturefile == ""} {return}
	set excelfile [file join [pwd] output report "test witness report.xlsx"]
	if {![file exists $excelfile]} {puts "report file not exist"; return}
	set pictureType [lindex [split [file tail $picturefile] "."] end]
	if {$pictureType eq "png"} {
		set pictureSaveFile $picturefile
	} else {set pictureSaveFile [file join [pwd] input "picture.png"]}
	if {![file exists $pictureSaveFile] && $pictureType eq "pdf"} {
		set pictureHandle [mupdf::open $picturefile]
		set pictureObj [$pictureHandle getpage 0]
		# 保存为图片
		$pictureObj savePNG $pictureSaveFile -zoom 4
		mupdf::close $pictureHandle
	}
	# 剪切图片
	set pictureNmaeList [croppicture $pictureSaveFile]
	try {
		set excel [::tcom::ref createobj "Excel.Application"]
		set workbooks [$excel Workbooks]
		set workbook [$workbooks Open $excelfile]
		set worksheets [$workbook Worksheets]
		set worksheet1 [$worksheets Item 1]
		set shape [$worksheet1 Shapes]
		set xlocation 100
		for {set i 0} {$i<[llength $pictureNmaeList]} {incr i} {
			set cropedImg [::cv::imread [lindex $pictureNmaeList $i]]
			set width [lindex [$cropedImg size] 1]
			puts "${width}::$xlocation"
			$shape AddPicture [lindex $pictureNmaeList $i] 0 1 $xlocation 100 -1 -1
			incr xlocation [expr $width-430]
			$cropedImg close
			 file delete [lindex $pictureNmaeList $i]
		}
		
	} on error {em} {
		puts $em
	}
	$workbook Save
	$workbook Close
	$excel Quit
	unset excel
	file delete $pictureSaveFile
}

# 剪切图片
proc croppicture {picturePath} {
	set picture [::cv::imread $picturePath $::cv::IMREAD_COLOR]
	# set pictureName [lindex [file tail $picture] 0]
	set width [lindex [$picture size] 1]
    set height [lindex [$picture size] 0]
	set pictureNmaeList {}
	if {$width >[expr $height*2]} {
		set num [expr int($width/$height)]
		set newWidth [expr int($width/$num)]
		for {set i 0} {$i<$num} {incr i} {
			set cropedPicture [$picture crop [expr $newWidth*$i] [expr int($height/6)] $newWidth [expr int($height*4/6)]]
			::cv::imwrite "[lindex [split $picturePath .] 0]_$i.png" $cropedPicture
			lappend pictureNmaeList "[lindex [split $picturePath .] 0]_$i.png"
			$cropedPicture close
		}
	}
	$picture close
	return $pictureNmaeList
}

# 更新图片
proc updateImg {v w} {
	if {[catch {$v isOpened}]} {puts "video off"; return;}
    if {[$v isOpened]} {
    	try {
            set f [$v read]
            if {$w ne ""} {$w write $f}

            set b_jpg [::cv::imencode "*.ppm" $f]
            set img [image create photo -data $b_jpg -format "ppm"]
            videoImg copy $img -subsample 1
            $f close
            image delete $img
        } on error {em} {
            puts $em
            catch {$w close}
            $v close
            videoImg blank
           	set ::run 0
        }
    }
}

# 循环调用
proc run {ms body} {
	global frameNum
	uplevel #0 $body
	incr frameNum 33
	if {$::run} {
		after $ms [info level 0]
	}
}

# 打开摄像头进行捕获图像
proc videoCapture {} {
    global width height img w v outVideo

    if {[info exist ::run] && $::run eq "1"} {return}
    # 打开摄像头
    set v [::cv::VideoCapture index 0]
    if {[$v isOpened]==0} {
        puts "Open camera $index failed."
        exit
    } else {
    	puts "Open camera succeeded"
    }

    # 获取图像的宽度和高度
    set width [$v get 3]
    set height [$v get 4]
    set name [clock format [clock second] -format %y-%m-%d-%H-%M-%S]
    file mkdir [file join [pwd] output videos]
    set outVideo [file join [pwd] output videos $name.avi]
    set w [::cv::VideoWriter $outVideo MJPG 30.0 $width $height 1]
    set ::run 1
    run 1 [list updateImg $v $w]
}

# 关闭摄像头
proc videoCaptureStop {} {
	global w v

	# stop record
	set ::run 0
    catch {$w close}
    catch {$v close}
    set v ""

    # clear image data
    videoImg blank
}



# 视频播放
proc videoPlay {} {
	global v frameNum playState videoPath

	# 正常播放完成后重新播放
	if {[info exist playState] && $playState eq "1"} {
		set videoPath [tk_getOpenFile]
		set frameNum 0
		if {$videoPath eq ""} {return}
		set v [::cv::VideoCapture file $videoPath]
	} elseif {$v == ""} {
		# 首次播放
		if {$videoPath eq ""} {
			set videoPath [tk_getOpenFile]
			set frameNum 0
			if {$videoPath eq ""} {return}
		}
		set v [::cv::VideoCapture file $videoPath]
	}

	# 设置视频的当前帧播放
	$v set $::cv::CAP_PROP_POS_MSEC $frameNum
	set playState 1
	while {$playState == 1 && [$v isOpened] == 1} {
		try {
			set f [$v read]
			set b_jpg [::cv::imencode "*.ppm" $f]
			set img [image create photo -data $b_jpg -format "ppm"]

	        #画布上显示图像
	        catch {videoImg copy $img -subsample 1}

			update
			$f close
			image delete $img
			set key [::cv::waitKey 10]
			if {$key==27} {
				break
			}
			incr frameNum 33
		} on error {em} {
            puts $em
            break
        }
	}
	videoImg blank
}

# 按记录跳转到指定视频位置
proc skip2mouse {skipTime} {
	global frameNum videoPath playState v

	set videotime [expr [$v get $::cv::CAP_PROP_FRAME_COUNT]*33]
	set videoname [lindex [split $videoPath "/"] end]
	set videoFileTime [expr [clock scan [lindex [split $videoname "."] 0] -format "%y-%m-%d-%H-%M-%S"] * 1000]
	set skipFrameNum [expr int(($skipTime-$videoFileTime)/33)]

	videoPause
	set frameNum [expr $skipFrameNum*33]
	videoPlay
}

# 视频暂停
proc videoPause {} {
	global v playState
	set playState 0
}

# 视频停止
proc videoStop {} {
	global v frameNum playState videoPath
	set frameNum 1

	# stop record
    set playState 0
    catch {$v close}
    set v ""
    set videoPath ""

    # clear image data
    videoImg blank
}

# 打开右上角的测试流程PDF文件查看器
proc openPDF {{filename ""}} {
	global PDFhandle pages
	if {$filename eq ""} {
		set filename [tk_getOpenFile]
		if {[string first "input" $filename] ne "-1"} {
			set time [clock format [clock second] -format %y-%m-%d-%H-%M]
			set name [file tail $filename]

		    file mkdir [file join [pwd] output report]
		    set outDoc [file join [pwd] output report $time-$name]
		    file copy -force $filename $outDoc
		    set filename $outDoc
		}
	}
	if {$filename eq ""} {return}
	if {[info exist PDFhandle]} {
		catch {$PDFhandle close}
	}
	set PDFhandle [mupdf::open $filename]
	set pages [$PDFhandle npages]
	mupdf::widget .main.frame.testframe.c $PDFhandle -bg darkgrey
	pack .main.frame.testframe.c -fill both -expand yes

	# 创建弹出菜单
    set popupMenu [menu .main.frame.testframe.c.popupMenu -tearoff 0]
    .main.frame.testframe.c.popupMenu add command -label "OK" -command { drawSymbol "OK" }
    .main.frame.testframe.c.popupMenu add command -label "Fail" -command { drawSymbol "Fail" }
    .main.frame.testframe.c.popupMenu add command -label "Clear" -command { clearSymbol }
    .main.frame.testframe.c.popupMenu add command -label "Play" -command { playSymbolVideo }
    .main.frame.testframe.c.popupMenu add command -label "Load picture to report" -command { loadPicture2excel }
    
    # 绑定鼠标右键事件
    bind .main.frame.testframe.c <ButtonPress-3> {
    	lassign [.main.frame.testframe.c win2page %x %y] x y
        tk_popup .main.frame.testframe.c.popupMenu %X %Y
    }
}

# 关闭右上角的测试流程PDF文件查看器
proc closePDF {} {
	global PDFhandle
	catch {set PDFhandle [mupdf::close $PDFhandle]}
	destroy .main.frame.testframe.c
}

# 打开左下角的站场图PDF文件查看器
proc openPlanPDF {{filename ""}} {
	global planPDFhandle
	if {$filename eq ""} {set filename [tk_getOpenFile]}
	if {$filename eq ""} {return}
	if {[info exist planPDFhandle]} {
		catch {$planPDFhandle close}
	}
	set planPDFhandle [mupdf::open $filename]
	set pages [$planPDFhandle npages]
	mupdf::widget .main.frame.cplan.c $planPDFhandle -bg darkgrey
	pack .main.frame.cplan.c -fill both -expand yes
}

# 关闭左下角的站场图PDF文件查看器
proc closePlanPDF {} {
	global planPDFhandle
	catch {set planPDFhandle [mupdf::close $planPDFhandle]}
	destroy .main.frame.cplan.c
}

# 在PDF查看器窗口中搜索文本
proc doSearch { pdfW } {
	global textSearch
	$pdfW search $textSearch		
}

# 在PDF查看器窗口中打开搜索面板
proc openSearchPanel { pdfW } {
	global textSearch
	
	set textSearch {}
	
	set panelW [toplevel $pdfW.searchPanel -padx 60 -pady 20]
	wm title $panelW "Search ..."
	wm attributes $panelW -topmost true
	 
	# don't use ttk::entry, since it cannot change the text color !
	entry $panelW.search -textvariable textSearch  
	button $panelW.ok -text "Search" 
	pack $panelW.search -fill x
	pack $panelW.ok

	$panelW.ok configure -command [list doSearch $pdfW]
	bind $panelW.search <Return> [list doSearch $pdfW] 

   	# place the new panel close to the pdfW widget
	set x0 [winfo rootx $pdfW]
	set y0 [winfo rooty $pdfW]
	wm geometry $panelW +[expr {$x0-10}]+[expr {$y0-10}] 

	# when this panel is closed, reset the search
	bind $panelW <Destroy> [list apply { 
		{W panelW pdfW} {
		     # NOTE: since <Destroy> is propagated to all children,
			 #  the following "if", ensure that this core script is executed
			 #  just once.
			if {  $W != $panelW } return
			$pdfW search ""
		}} %W $panelW $pdfW]
		
}

# 在右上角的测试流程图中画“√”或者“×”
proc drawSymbol {var} {
	global PDFhandle x y
	variable currentPage

	set pageHandle [$PDFhandle getpage $currentPage]

	if {$var eq "OK"} {
		for {set i -7} {$i <= 0} {incr i} {
		    set x1 [expr $x+$i-2]
		    set y1 [expr $y+$i-2]
		    set x2 [expr $x1+2]
		    set y2 [expr $y1+2]
		    lappend annotationID [$pageHandle annot create highlight -color {0 1 0} -vertices "$x1 $y1 $x2 $y2"]
		}

		for {set i 0} {$i <= 15} {incr i} {
		    set x3 [expr $x+$i-2]
		    set y3 [expr $y-$i-2]
		    set x4 [expr $x3+2]
		    set y4 [expr $y3+2]
		    lappend annotationID [$pageHandle annot create highlight -color {0 1 0} -vertices "$x3 $y3 $x4 $y4"]
		}
	} else {
		for {set i -7} {$i <= 7} {incr i} {
		    set x1 [expr $x+$i-2]
		    set y1 [expr $y+$i-2]
		    set x2 [expr $x1+2]
		    set y2 [expr $y1+2]
		    lappend annotationID [$pageHandle annot create highlight -color {1 0 0} -vertices "$x1 $y1 $x2 $y2"]
		}

		for {set i -7} {$i <= 7} {incr i} {
		    set x3 [expr $x+$i-2]
		    set y3 [expr $y-$i-2]
		    set x4 [expr $x3+2]
		    set y4 [expr $y3+2]
		    lappend annotationID [$pageHandle annot create highlight -color {1 0 0} -vertices "$x3 $y3 $x4 $y4"]
		}
	}
	catch {.main.frame.testframe.c refreshpage}; # with modified version of muPdfWidget
	# 坐标和ID存放到字典中
	set annotIDLocation [dict create]
	dict set annotIDLocation "$x,$y-$currentPage ID" $annotationID
	dict set annotIDLocation "$x,$y-$currentPage Time" [clock milliseconds]

	# 将数据写入文件
	set filename [file join [pwd] output data annotIDLocation.txt]
	set file [open $filename "a"]
	puts $file $annotIDLocation
	close $file
}

# 在右上角的测试流程图中清除“√”或者“×”
proc clearSymbol {} {
	global PDFhandle x y
	variable currentPage

	set pageHandle [$PDFhandle getpage $currentPage]
	set filename [file join [pwd] output data annotIDLocation.txt]
	set file [open $filename r]
	set fileData [read $file]
	close $file

	dict for {key value} $fileData {
		set coords_page [split [lindex $key 0] "-"]
		if {[lindex $key end] eq "ID" && [lindex $coords_page end] eq $currentPage} {
			set coords [split [lindex $coords_page 0] ","]
			# 计算两点之间的距离
			set distance [math::geometry::distance $coords "$x $y"]
			if {$distance < 60} {
				if {![info exist value] || [llength $value] < 1} {
					continue
				}

				# 删除注释
				foreach annotID $value {
					$pageHandle annot $annotID delete
				}

				# 清空文本的记录
				dict unset fileData $key
				dict unset fileData [lreplace $key end end Time]
				set file [open $filename w]
				puts $file $fileData
				close $file
				break
			}
		}
	}
	catch {.main.frame.testframe.c refreshpage}; # with modified version of muPdfWidget
}

# 根据标记结果进行回放视频
proc playSymbolVideo {} {
	global PDFhandle x y
	variable currentPage

	set pageHandle [$PDFhandle getpage $currentPage]
	set filename [file join [pwd] output data annotIDLocation.txt]
	set file [open $filename r]
	set fileData [read $file]
	close $file

	dict for {key value} $fileData {
		set coords_page [split [lindex $key 0] "-"]
		if {[lindex $key end] eq "Time" && [lindex $coords_page end] eq $currentPage} {
			set coords [split [lindex $coords_page 0] ","]
			# 计算两点之间的距离
			set distance [math::geometry::distance $coords "$x $y"]
			if {$distance < 60} {
				if {![info exist value] || [llength $value] < 1} {
					continue
				}

				skip2mouse $value
				break
			}
		}
	}
}

# 更新评论区信息
proc updatecomments {commentsarear} {
	global commentsPath
	file mkdir [file join [pwd] output data]
	set commentsPath [file join [pwd] output data comment.txt]
	if {$commentsPath == ""} {return}
	set f [open $commentsPath "r"]
	fconfigure $f -encoding utf-8
	if {[winfo class $commentsarear] == "Text"} {
		$commentsarear delete 1.0 end 
		while {[gets $f line] != -1} {
			if {$line == ""} {continue}
			$commentsarear insert end "$line\n"
		}
	} else {
		$commentsarear delete 0 end
		while {[gets $f line] != -1} {
			if {$line == ""} {continue}
			$commentsarear insert end "$line"
		}
		$commentsarear see end
	}
	close $f
}

# 打开提交评论的窗口
proc commentSave {} {
	global mainframe commentsPath
    
	if {![winfo exists .comments]} {
		toplevel .comments
		wm title .comments "comments"
		.comments configure -background "#ffffff"
		text .comments.text -height 10 -width 40 -setgrid true -font {Helvetica 12} -wrap word
		set buttonarea [frame .comments.button -background "#ffffff"]
		ttk::button $buttonarea.commit -text "commit" -command [list commit $commentsPath] -width 10
		grid .comments.text -sticky news
		pack $buttonarea.commit -side right 
		grid $buttonarea -row 1 -column 0 -columnspan 2 -sticky news -padx 20 -pady 8
		grid rowconfigure .comments 0 -weight 1
		grid columnconfigure .comments 0 -weight 1
	} else {
		wm deiconify .comments
	}
}

# 提交评论
proc commit {commentsPath} {
	global commentsarear
    set commentsdata [.comments.text get 1.0 end-1c ]
    set time [clock format [clock second] -format %y-%m-%d-%H-%M-%S]
    if {$commentsdata ne ""} {
        # 插入评论和时间
		set fileHandle [open $commentsPath "a"]

		# 设置文件句柄的编码为UTF-8
		fconfigure $fileHandle -encoding utf-8

		# 写入文本内容
		if {[llength $commentsdata] == 1} {puts $fileHandle "$time:$commentsdata"}
		if {[llength $commentsdata] > 1} {
			set strdata ""
			set num 1
			foreach data $commentsdata {
				if {$num == [llength $commentsdata]} {
					append strdata "${data}."
					break
				}
				append strdata "${data},"
				incr num
			}
			puts $fileHandle "$time: $strdata"
		}
		
		# 关闭文件
		close $fileHandle
		.comments.text delete 1.0 end
		updatecomments $commentsarear
    }
	wm withdraw .comments
}

# 通过菜单打开更改评论的窗口
proc alterComments {} {
	global commentsPath
	if {[winfo exists .alterComments]} {
		wm deiconify .alterComments
		updatecomments .alterComments.text
	} else {
		toplevel .alterComments
		wm title .alterComments "alterComments"
		.alterComments configure -background "#ffffff"
		text .alterComments.text -height 10 -width 40 -setgrid true -font {Helvetica 12} -wrap word
		set buttonarea [frame .alterComments.button -background "#ffffff"]
		ttk::button $buttonarea.commit -text "commit" -command [list alterComment $commentsPath] -width 10
		grid .alterComments.text -sticky news
		pack $buttonarea.commit -side right
		grid $buttonarea -row 1 -column 0 -sticky news -columnspan 2 -padx 20 -pady 8
		grid rowconfigure .alterComments 0 -weight 1
		grid columnconfigure .alterComments 0 -weight 1
		updatecomments .alterComments.text
	}
}

# 更改评论
proc alterComment {commentsPath} {
	global commentsarear
	set commentsdata [.alterComments.text get 1.0 end-1c]

	set f [open $commentsPath w]
	fconfigure $f -encoding utf-8

	puts $f $commentsdata
	close $f
	updatecomments $commentsarear
	
	wm withdraw .alterComments
}

proc readReportTemplate {} {
	global allTitles
	set allTitles ""

	set app [Excel OpenNew false]
	set filename [file join [pwd] input "test witness report.xlsx"]
	set wb [Excel OpenWorkbook $app $filename]
	set ws [Excel GetWorksheetIdByName $wb "Test case summarised"]
	set matrix [Excel GetWorksheetAsMatrix $ws]

	foreach line [lrange $matrix 2 end] {
		# set value [Excel GetCellValue $ws($id) $row $colres]
		if {[lindex $line 2] ne ""} {
			lappend allTitles [lindex $line 2]
		}
	}

	#Close Workbook
    Excel Close $wb

	#close Excel
	Excel Quit $app 0
}

# 储存功能点标题所在的页面及坐标
proc saveTitlePos {} {
	global titleDataRange PDFhandle pages allTitles
	set titlePageData ""
	set titleDataRange [dict create]

	readReportTemplate
	foreach title $allTitles {
		foreach page_coords [$PDFhandle search $title] {
			set page [lindex $page_coords 0]
			if {$page > 9} {
				lassign [lindex $page_coords end] x1 y1 x2 y2
				break
			}
		}

		if {$page > 9} {
			lappend titlePageData "{$title} $page $y1"
		}
	}

	for {set i 0} {$i < [llength $titlePageData]} {incr i} {
		set data1 [lindex $titlePageData $i]
		if {[expr $i + 1] < [llength $titlePageData]} {
			set data2 [lindex $titlePageData [expr $i + 1]]
			set pageRange "[lindex $data1 1]-[lindex $data2 1]"
			set coordsRange "[lindex $data1 2]-[lindex $data2 2]"
		} else {
			set pageRange "[lindex $data1 1]-$pages"
			set coordsRange "[lindex $data1 2]-1080"
		}
		
		dict set titleDataRange [lindex $data1 0] [list $pageRange $coordsRange]
	}
}

# 每一个测试功能点的结果输出
proc getFunctionResult {text} {
	global titleDataRange

	set filename [file join [pwd] output data annotIDLocation.txt]
	set file [open $filename r]
	fconfigure $file -encoding utf-8
	set fileData [read $file]
	close $file
	set result ""

	foreach title [dict keys $titleDataRange] {
		if {$title eq $text} {
			set pageRange [lindex [dict get $titleDataRange $title] 0]
			lassign [split $pageRange "-"] page1 page2
			
			set coordsRange [lindex [dict get $titleDataRange $title] 1]
			lassign [split $coordsRange "-"] coords1 coords2

			dict for {key value} $fileData {
				set coords_page [split [lindex $key 0] "-"]
				if {[lindex $coords_page end] <= $page2 && [lindex $coords_page end] >= $page1} {
					set coords [split [lindex $coords_page 0] ","]

					if {[lindex $coords_page end] == $page1 && [lindex $coords 1] > $coords1 || [lindex $coords_page end] == $page2 && [lindex $coords 1] < $coords2} {
						if {[llength $value] eq "30"} {
							lappend result "\u00D7,[expr [lindex $coords_page end] + 1]"
						} elseif {[llength $value] eq "24"} {
							lappend result "\u221A"
						}
					}
				}
			}

			if {[string first "\u00D7" $result] eq "-1"} {
				if {[string first "\u221A" $result] ne "-1"} {
					set result "\u221A"
				} else {
					set result "-"
				}
			}
		}
	}

	return $result
}

proc fillReport {} {
	global titleDataRange
	
	saveTitlePos
	set app [Excel OpenNew false]
	set filename [file join [pwd] input "test witness report.xlsx"]
	set name [file tail $filename]
	set time [clock format [clock second] -format %y-%m-%d-%H-%M]
    set outDoc [file join [pwd] output report $time-$name]
    file copy -force $filename $outDoc


	set wb [Excel OpenWorkbook $app $outDoc]
	set ws [Excel GetWorksheetIdByName $wb "Test case summarised"]
	set matrix [Excel GetWorksheetAsMatrix $ws]


	set row 3
	foreach {key value} $titleDataRange {
		set res [getFunctionResult $key]
		Excel SetCellValue $ws $row 4 $res
		incr row
	}

	#save & close workbooks
    Excel SaveAs $wb $outDoc
    Excel Close $wb
    # close Excel
    Excel Quit $app 0

    # force quit
    set pids [concat [twapi::get_process_ids -name EXCEL.exe] [twapi::get_process_ids -name wps.exe]]
    #end process by id
    foreach pid $pids {
        twapi::end_process $pid -force
    }
    tk_messageBox -type ok -title "Generate Report" -message "Report generated successfully"
}

console show
buildGUI