Function GetObject() doesn't handle errors well.  If the object isn't found, it throws a generic error, making debugging difficult.  It needs more specific error handling to indicate *why* the object retrieval failed.