<%
function  week_cstr(dtDate)
	select case cint(Weekday(dtDate))
            case 1
		strwk="日"
	case 2
		strwk="一"
	case 3
		strwk="二"
	case 4
		strwk="三"
	case 5
		strwk="四"
	case 6
		strwk="五"
	case 7
		strwk="六"
	case else
	end select
    week_cstr=strwk
end function			
%>			