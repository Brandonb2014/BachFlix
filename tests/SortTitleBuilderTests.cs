using System;
using BachFlixNfo.Features;

internal static class SortTitleBuilderTests
{
    private static void Expect(string input, string expected)
    {
        string actual = SortTitleBuilder.Build(input);
        if (actual != expected)
            throw new Exception("Build(\"" + input + "\") returned \"" + actual + "\", expected \"" + expected + "\"");
    }

    public static int Main()
    {
        try
        {
            Expect("1 Night in San Diego (2020)", "00001 Night in San Diego (2020)");
            Expect("13 Moons (2002)", "00013 Moons (2002)");
            Expect("1780 (2025)", "01780 (2025)");
            Expect("10000 BC (2008)", "10000 BC (2008)");
            Expect("2001: A Space Odyssey (1968)", "02001: A Space Odyssey (1968)");
            Expect("02000 Mules (2022)", "02000 Mules (2022)");
            Expect("123456 Title (2000)", "123456 Title (2000)");
            Expect("The 13th Warrior (1999)", "00013th Warrior (1999)");
            Expect("The Matrix (1999)", "Matrix (1999)");
            Expect("Up (2009)", "Up (2009)");
            Expect("Ocean's 11 (2001)", "Ocean's 11 (2001)");
            Expect("", "");

            Console.WriteLine("Sort title builder tests passed.");
            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine(ex.Message);
            return 1;
        }
    }
}
