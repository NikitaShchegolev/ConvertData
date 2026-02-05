namespace ConvertData.Domain;

/// <summary>
/// Доменная модель строки таблицы.
/// Представляет одну запись с параметрами профиля/кодов и числовыми полями.
/// </summary>
internal sealed class Row
{
    /// <summary>Тип/категория (например "Н2").</summary>
    public string Name { get; set; } = "";

    /// <summary>Код соединения/элемента (например "H2_01").</summary>
    public string CONNECTION_CODE { get; set; } = "";

    /// <summary>Профиль/марка (например "16Б1").</summary>
    public string Profile { get; set; } = "";

    /// <summary>Числовое поле Nt.</summary>
    public int Nt { get; set; }

    /// <summary>Числовое поле Nc.</summary>
    public int Nc { get; set; }

    /// <summary>Числовое поле N.</summary>
    public int N { get; set; }

    /// <summary>Числовое поле Qo.</summary>
    public int Qo { get; set; }

    /// <summary>Числовое поле Q.</summary>
    public int Q { get; set; }

    /// <summary>Числовое поле T.</summary>
    public int T { get; set; }

    /// <summary>Числовое поле M.</summary>
    public int M { get; set; }

    /// <summary>Числовое поле Mneg.</summary>
    public double Mneg { get; set; }

    /// <summary>Числовое поле Mo.</summary>
    public double Mo { get; set; }

    /// <summary>Поле α (альфа) — хранится под ASCII-именем, но сериализуется в ключ "α".</summary>
    public double Alpha { get; set; }

    /// <summary>Поле β (бета) — хранится под ASCII-именем, но сериализуется в ключ "β".</summary>
    public double Beta { get; set; }

    /// <summary>Поле γ (гамма) — хранится под ASCII-именем, но сериализуется в ключ "γ".</summary>
    public double Gamma { get; set; }

    /// <summary>Поле δ (дельта) — хранится под ASCII-именем, но сериализуется в ключ "δ".</summary>
    public double Delta { get; set; }

    /// <summary>Поле ε (эпсилон) — хранится под ASCII-именем, но сериализуется в ключ "ε".</summary>
    public double Epsilon { get; set; }

    /// <summary>Поле λ (лямбда) — хранится под ASCII-именем, но сериализуется в ключ "λ".</summary>
    public double Lambda { get; set; }
}
