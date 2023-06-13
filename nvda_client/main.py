import sys
from pathlib import Path
import ctypes
import time


class NVDAClient:
    """A class for interacting with NVDA using the NVDA Controller Client.

    Methods:
        speak: Speak the given text using NVDA.
        braille: Display the given text on a connected braille display.
        braille_and_speak: Display the given text on a connected braille
            display and speak it.
        cancelSpeech: Cancel any speech currently being spoken by NVDA.
    """

    def __init__(self) -> None:
        """Initialize the NVDAClient class.

        Args:
            None

        Returns:
            None

        Raises:
            RuntimeError: If the NVDA Controller Client is not running.
        """

        if sys.maxsize > 2**32:  # 64-bit
            dll_path = (
                Path(__file__).parent.absolute() / "nvdaControllerClient64.dll"
            )
        else:  # 32-bit
            dll_path = (
                Path(__file__).parent.absolute() / "nvdaControllerClient32.dll"
            )
        self._nvdaControllerClient = ctypes.windll.LoadLibrary(str(dll_path))

        if (
            result := self._nvdaControllerClient.nvdaController_testIfRunning()
        ) != 0:
            raise RuntimeError(self._get_error_message(result))

    @staticmethod
    def _get_error_message(error_code: int) -> str:
        """Get the error message for a given error code.

        Args:
            error_code (int): The error code to get the message for.

        Returns:
            str: The error message.

        Raises:
            None
        """
        return str(ctypes.WinError(error_code))

    def speak(self, text: str) -> None:
        """Speak the given text using NVDA.

        Args:
            text (str): The text to speak.

        Returns:
            None

        Raises:
            RuntimeError: If the NVDA Controller Client returns an error code.
        """
        if (
            result := self._nvdaControllerClient.nvdaController_speakText(text)
        ) != 0:
            raise RuntimeError(self._get_error_message(result))

    def braille(self, text: str) -> None:
        """Display the given text on a connected braille display.

        Args:
            text (str): The text to display.

        Returns:
            None

        Raises:
            RuntimeError: If the NVDA Controller Client returns an error code.
        """
        if (
            result := self._nvdaControllerClient.nvdaController_brailleMessage(
                text
            )
        ) != 0:
            raise RuntimeError(self._get_error_message(result))

    def braille_and_speak(self, text: str) -> None:
        """Display the given text on a connected braille display and speak it.

        Args:
            text (str): The text to display and speak.

        Returns:
            None

        Raises:
            RuntimeError: If the NVDA Controller Client returns an error code.
        """
        self.braille(text)
        self.speak(text)

    def cancelSpeech(self) -> None:
        """Cancel any speech currently being spoken by NVDA.

        Args:
            None

        Returns:
            None

        Raises:
            RuntimeError: If the NVDA Controller Client returns an error code.
        """
        if (
            result := self._nvdaControllerClient.nvdaController_cancelSpeech()
        ) != 0:
            raise RuntimeError(self._get_error_message(result))


if __name__ == "__main__":
    pass
