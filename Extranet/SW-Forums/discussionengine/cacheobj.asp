<SCRIPT LANGUAGE="JavaScript" RUNAT="Server">

// ======================================================================
//
// CACHE OBJECT
//
// ======================================================================

function SWEFCache
(
)
{
  return this;
}

// ======================================================================
//
// Interface to private member variables.
//
// ======================================================================

// ======================================================================
//
// Main object methods.
//
// ======================================================================

function retrieveObjectByKey_cch_disc
(
 strKey
)
{
  var objCachedObject = undefined_disc;
  if (config.CACHE_Enabled)
    {
      strKey = this.unmangleKey (strKey);

      var strLastUpdated = Application (strKey + config.CACHE_TimeStampKey);
      if (isDefined_disc (strLastUpdated))
	{
	  if (this.cacheObjectAlive (strLastUpdated))
	    {
	      objCachedObject = Application (strKey);
	    }
	}
    }

  return objCachedObject;
}

function storeObjectByKey_cch_disc
(
 strKey,
 objCachedObject
)
{
  if (config.CACHE_Enabled)
    {
      strKey = this.unmangleKey (strKey);

      Application.Lock ();
      Application (strKey) = objCachedObject;
      Application (strKey + config.CACHE_TimeStampKey) = (new Date ()).toLocaleString ();
      Application.Unlock ();
    }

  return;
}

function removeObjectByKey_cch_disc
(
 strKey
)
{
  if (config.CACHE_Enabled)
    {
      strKey = this.unmangleKey (strKey);

      this.removeObjectPrivateRoutine (strKey);
      this.removeObjectPrivateRoutine (strKey + config.CACHE_TimeStampKey);
    }

  return;
}

function removeObjectPrivateRoutine_cch_disc
(
 strUnmangledkey
)
{
  Application.Lock ();
  if (isDefined_disc (Application (strUnmangledkey)))
    {
      if (typeof (Application (strUnmangledkey)) == "object")
	{
	  Application (strUnmangledkey).close ();
	}
      Application (strUnmangledkey) = undefined_disc;
    }
  Application.Unlock ();

  return;
}

function removeObjectsByKeyPrefix_cch_disc
(
 strKeyPrefix
)
{
  strKeyPrefix = this.unmangleKey (strKeyPrefix);
  var strCurrentKey;
  for (var nCounter = 1; nCounter <= Application.Contents.Count; nCounter++)
    {
      strCurrentKey = Application.Contents.Key (nCounter);
      if (strKeyPrefix.toUpperCase () == (strCurrentKey.substring (0, strKeyPrefix.length)).toUpperCase ())
	{
	  this.removeObjectPrivateRoutine (strCurrentKey);
	}
    }

  return;
}

function purge_cch_disc
(
)
{
  Response.Write ("<H2>Purging</H2>");
  this.removeObjectsByKeyPrefix (config.getUniqueKey ());

  return;
}

// ======================================================================
//
// Private object methods.  These should not be called except by this object.
//
// ======================================================================

function unmangleKey_cch_disc
(
 strKey
)
{
  return config.getUniqueKey () + strKey;
}

function cacheObjectAlive_cch_disc
(
 strUpdateTime
)
{
  var bAreWeAlive;
  var dtItemCachedAt = new Date (strUpdateTime);
  var dtCurrentTimeStamp = new Date ();
  var nCacheTimeout = config.ADMINSETTING_CacheTimeoutMinutes * 60 * 1000;

  if (((dtCurrentTimeStamp - dtItemCachedAt) < nCacheTimeout)
      || (dtItemCachedAt.getDate () != dtCurrentTimeStamp.getDate ()))
    {
      bAreWeAlive = true;
    }
  else
    {
      bAreWeAlive = false;
    }

  delete dtItemCachedAt;
  delete dtCurrentTimeStamp;
  return bAreWeAlive;
}

SWEFCache.purge = purge_cch_disc;

SWEFCache.prototype.retrieveObjectByKey = retrieveObjectByKey_cch_disc;
SWEFCache.prototype.storeObjectByKey = storeObjectByKey_cch_disc;
SWEFCache.prototype.removeObjectByKey = removeObjectByKey_cch_disc;
SWEFCache.prototype.removeObjectPrivateRoutine = removeObjectPrivateRoutine_cch_disc;
SWEFCache.prototype.removeObjectsByKeyPrefix = removeObjectsByKeyPrefix_cch_disc;
SWEFCache.prototype.purge = purge_cch_disc;
SWEFCache.prototype.unmangleKey = unmangleKey_cch_disc;
SWEFCache.prototype.cacheObjectAlive = cacheObjectAlive_cch_disc;
</SCRIPT>

