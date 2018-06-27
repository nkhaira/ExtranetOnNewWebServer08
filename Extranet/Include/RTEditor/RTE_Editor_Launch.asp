<SCRIPT Language = "JavaScript">
  function RTEditor_Open(Form ,Element, Site_ID, Site_Code, Length, Cols, Rows) {
    //alert("<%=Translate("IMPORTANT: Remember to SAVE your changes on this form after you exit from the HTML Editor, otherwise your changes will be lost.",Login_Language,conn)%>");
    var RTE_Form      = Form;
    var RTE_Element   = Element;
    var RTE_Site_ID   = Site_ID;
    var RTE_Site_Code = Site_Code;
    var RTE_Length    = Length;
    var RTE_Cols      = Cols;
    var RTE_Rows      = Rows;
    var URL = "/SW-Administrator/SW-RTEditor.asp?Form=" + RTE_Form + "&Element=" + RTE_Element + "&Site_ID=" + RTE_Site_ID + "&Site_Code=" + RTE_Site_Code + "&Length=" + RTE_Length + "&Cols=" + RTE_Cols + "&Rows=" + RTE_Rows;
    var RTEditor = window.open(URL,null,"height=480,width=800,status=no,toolbar=no,resizable=yes,menubar=no,location=no");
    RTEditor.window.focus();
  }
</SCRIPT>
