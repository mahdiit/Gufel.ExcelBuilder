﻿namespace Gufel.ExcelBuilder.Model.Base;

public class ExcelBuildException(string msg, string? code = null) : Exception(msg)
{
    public string? ErrorCode { get; private set; } = code;
}