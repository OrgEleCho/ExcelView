using System.Text;

public static class Utils
{
    public static IEnumerable<string> EnumerateLetters()
    {
        List<int> digits = new() { 0 };

        StringBuilder sb = new();

        while (true)
        {
            digits[0]++;

            for (int i = 0; i < digits.Count; i++)
            {
                if (digits[i] >= 27)
                {
                    digits[i] -= 26;
                    if (i == digits.Count - 1)
                        digits.Add(0);

                    digits[i + 1]++;
                }
            }

            sb.Clear();
            for (int i = digits.Count - 1; i >= 0; i--)
            {
                if (digits[i] == 0)
                    continue;

                sb.Append((char)('A' + digits[i] - 1));
            }

            yield return sb.ToString();
        }
    }
}