﻿namespace OOXMLValidatorCLI.Interfaces
{
    public interface IValidate
    {
        object OOXML(string filePath, string format, bool returnXml, bool recursive, bool includeValid);
    }
}
