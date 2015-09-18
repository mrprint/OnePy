namespace Enterprise.AddIn
{
    public enum COMError: uint
    {
        E_UNEXPECTED = 0x8000FFFF,
        E_POINTER = 0x80004003,
        E_FAIL = 0x80004005,
        S_FALSE = 1,
        S_OK = 0
    }
}
