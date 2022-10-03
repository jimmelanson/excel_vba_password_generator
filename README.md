# excel_vba_password_generator
A utility using rand() and rank() to generate passwords that comply with multiple rules.

This password generator was written because a.) I wanted to solve some string manipulation problems and b.) I was bored.

This example workbook generates passwords that follow these rules/contraints:

- It uses upper case letters.
- It uses lower case letters.
- It uses digits.
- It uses special characters (punctuation).
- It lets you set a minimum number of total characters.
- It lets you set a minimum number of each character type  (upper, lower, digits, special).
- It will not print dictionary words (other than brr and hmm and the like)
- It will not print two contiguous characters of the same type.
- It will not duplicate any character in the password.

While the idea of a password generator may seem straightforward, it actually is straightforward. But it’s not quick. There were several hurdles I had to overcome that made me backtrack a few times and even delete code to rewrite it from scratch. In the end, I was surprised at how simple it was overall. In total, I spent about 14 hours creating this. There are lots of lessons you can take from the code in this.

I think the thing that I learned the most from was combining Excel VBAs rand() and rank() procedures to come up with truly randomized character sets (passwords).

You can also take the library file and import it to your own project if you want to add a password generator to it.
