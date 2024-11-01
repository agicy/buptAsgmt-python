"""
This module provides functions to
    - count word occurrences in text,
    - merge word counts from two texts,
    - count words by their initial letters.
It also includes a main function to demonstrate these capabilities using two input text files.

Functions:
    count_word_occurrences(text: str) -> dict[str, int]:
        Counts the occurrences of each word in the provided text.

    merge_word_counts(dict1: dict[str, int], dict2: dict[str, int]) -> dict[str, int]:
        Merges two word count dictionaries by summing the counts of common words.

    count_words_by_initial(word_count_dict: dict[str, int]) -> dict[str, int]:
        Counts the total occurrences of words starting with each letter of the alphabet.

    main(article_path1, article_path2) -> None:
        The main function that reads two text files, cleans the text, counts word occurrences,
        merges the counts, and counts words by initial letters.

"""

import sys


def count_word_occurrences(text: str) -> dict[str, int]:
    """
    Counts the occurrences of each word in the provided text.

    Args:
        text (str): The text to count word occurrences in.

    Returns:
        dict[str, int]: A dictionary with words as keys and their counts as values.
    """
    words: list[str] = text.split()
    return {word: words.count(word) for word in words}


def merge_word_counts(dict1: dict[str, int], dict2: dict[str, int]) -> dict[str, int]:
    """
    Merges two word count dictionaries by summing the counts of common words.

    Args:
        dict1 (dict[str, int]): The first word count dictionary.
        dict2 (dict[str, int]): The second word count dictionary.

    Returns:
        dict[str, int]: A merged dictionary with summed word counts.
    """
    return {
        key: dict1.get(key, 0) + dict2.get(key, 0) for key in set(dict1) | set(dict2)
    }


def count_words_by_initial(word_count_dict: dict[str, int]) -> dict[str, int]:
    """
    Counts the total occurrences of words starting with each letter of the alphabet.

    Args:
        word_count_dict (dict[str, int]):
            A dictionary with words as keys and their counts as values.

    Returns:
        dict[str, int]: A dictionary with initial letters as keys and total counts as values.
    """
    return {
        initial: sum(
            count
            for word, count in word_count_dict.items()
            if word[0].lower() == initial
        )
        for initial in {chr(code) for code in range(ord("a"), ord("z") + 1)}
    }


def main(article_path1, article_path2) -> None:
    """
    The main function that reads two text files, cleans the text, counts word occurrences,
    merges the counts, and counts words by initial letters.

    Args:
        article_path1 (str): The path to the first text file.
        article_path2 (str): The path to the second text file.

    Returns:
        None
    """
    try:
        with open(file=article_path1, mode="r", encoding="utf8") as file1:
            s1: str = file1.read()

        with open(file=article_path2, mode="r", encoding="utf8") as file2:
            s2: str = file2.read()

        s1: str = "".join(
            ch if ch.isalnum() or ch.isspace() else " " for ch in s1
        ).lower()
        s2: str = "".join(
            ch if ch.isalnum() or ch.isspace() else " " for ch in s2
        ).lower()

        # Function calls and print statements as per the assignment
        d1: dict[str, int] = count_word_occurrences(text=s1)
        print(f"Total words in s1: {len(s1.split())}")
        print(f"Unique words in d1: {len(d1)}")

        d2: dict[str, int] = count_word_occurrences(text=s2)
        print(f"Total words in s2: {len(s2.split())}")
        print(f"Unique words in d2: {len(d2)}")
        print("")

        d3: dict[str, int] = merge_word_counts(dict1=d1, dict2=d2)
        print(f"Unique words in d3: {len(d3)}")
        print("Word counts in d3:")
        for word, count in d3.items():
            print(f"\t{word}: {count}")
        print("")

        d4: dict[str, int] = count_words_by_initial(word_count_dict=d3)
        print("Initial letter counts in d4:")
        for initial, count in d4.items():
            print(f"\t{initial}: {count}")
        print("")

    except FileNotFoundError:
        print("One of the provided file paths does not exist.")


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Please provide two file paths as command line arguments.")
    else:
        main(article_path1=sys.argv[1], article_path2=sys.argv[2])
