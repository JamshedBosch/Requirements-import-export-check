class HelperFunctions:

    @staticmethod
    def normalize_text(text, ignore_spaces_and_semicolons=True):
        """
        Normalize the given text by removing spaces, semicolons, quotes, and tab characters.

        :param text: The text to normalize.
        :param ignore_spaces_and_semicolons: Whether to remove spaces and semicolons.
        :return: The normalized text.
        """
        if not isinstance(text, str):  # Ensure text is a string
            text = str(text)  # Convert it to a string if it's not

        if ignore_spaces_and_semicolons:
            text = text.replace(" ", "")  # Remove spaces
            text = text.replace("\t", "")  # Remove tab characters
            text = text.replace(";", "")  # Remove semicolons
            text = text.replace("'", "")  # Remove single quotes
            text = text.replace('"', "")  # Remove double quotes

        # Normalize all types of whitespace (e.g., newlines, extra spaces)
        text = text.strip()  # Remove leading/trailing whitespace

        return text