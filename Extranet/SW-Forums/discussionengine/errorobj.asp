<SCRIPT LANGUAGE="JavaScript" RUNAT="Server">

// ======================================================================
//
// ERROR OBJECT
//
// ======================================================================

function SWEFError
(
 nErrorNumber,
 strDescription
)
{
  this.number = nErrorNumber;
  this.description = strDescription;

  return this;
}

</SCRIPT>