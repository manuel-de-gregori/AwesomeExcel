﻿using AwesomeExcel.Common.Models;

namespace AwesomeExcel.Customization.Models;

public class FontStyleCustomization { }

public class FontStyleCustomization<T> : FontStyleCustomization
{
    public Func<T, string> Name { get; set; }
    public Func<T, Color?> Color { get; set; }
    public Func<T, short?> HeightInPoints { get; set; }
    public Func<T, bool?> IsBold { get; set; }
}
