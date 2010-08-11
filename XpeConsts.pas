unit XpeConsts;

interface

const
  RESOURCETYPE_FILE      = 1;
  RESOURCETYPE_REGISTRY  = 2;

const
  PROPFORMAT_BINARY      = 1;
  PROPFORMAT_SZ          = 2;
  PROPFORMAT_INTEGER     = 3;
  PROPFORMAT_BOOL        = 4;
  PROPFORMAT_MULTISZ     = 5;
  PROPFORMAT_EXPRESSION  = 6;
  PROPFORMAT_GUID        = 9;


const
  REGCOND_ALWAYS_WRITE          = 1;  {Always write}
  REGCOND_WRITE_ONLY_IF_EXISTS  = 2;  {Only if value already exists}
  REGCOND_WRITE_ONLY_IF_MISSING = 3;  {Only if doesn't exist}
  REGCOND_DELETE                = 4;  {Delete the key or value}
  REGCOND_ALWAYS_EDIT           = 5;  {Always edit}
  REGCOND_EDIT_IF_EXISTS        = 6;  {Edit only if value exists}


implementation

end.
