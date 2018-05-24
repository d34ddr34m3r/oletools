"""
log_helper.py

General logging helpers

.. codeauthor:: Intra2net AG <info@intra2net>
"""

# === LICENSE =================================================================
#
# Redistribution and use in source and binary forms, with or without
# modification, are permitted provided that the following conditions are met:
#
#  * Redistributions of source code must retain the above copyright notice,
#    this list of conditions and the following disclaimer.
#  * Redistributions in binary form must reproduce the above copyright notice,
#    this list of conditions and the following disclaimer in the documentation
#    and/or other materials provided with the distribution.
#
# THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
# AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
# IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE
# ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE
# LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR
# CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF
# SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS
# INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN
# CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE)
# ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE
# POSSIBILITY OF SUCH DAMAGE.

# -----------------------------------------------------------------------------
# CHANGELOG:
# 2017-12-07 v0.01 CH: - first version
# 2018-02-05 v0.02 SA: - fixed log level selection and reformatted code
# 2018-02-06 v0.03 SA: - refactored code to deal with NullHandlers
# 2018-02-07 v0.04 SA: - fixed control of handlers propagation
# 2018-04-23 v0.05 SA: - refactored the whole logger to use an OOP approach

# -----------------------------------------------------------------------------
# TODO:


from __future__ import print_function
from ._json_formatter import JsonFormatter
from ._logger_class import OletoolsLogger
from . import _root_logger_wrapper
import logging
import sys


LOG_LEVELS = {
    'debug': logging.DEBUG,
    'info': logging.INFO,
    'warning': logging.WARNING,
    'error': logging.ERROR,
    'critical': logging.CRITICAL
}

DEFAULT_LOGGER_NAME = 'oletools'
DEFAULT_MESSAGE_FORMAT = '%(levelname)-8s %(message)s'


class LogHelper:
    def __init__(self):
        self._all_names = set()  # set so we do not have duplicates
        self._use_json = False
        self._is_enabled = False
        self._json_formatter = JsonFormatter()
        logging.setLoggerClass(OletoolsLogger)

    def get_or_create_silent_logger(self, name=DEFAULT_LOGGER_NAME, level=logging.CRITICAL + 1):
        """
        Get a logger or create one if it doesn't exist, setting a NullHandler
        as the handler (to avoid printing to the console).
        By default we also use a higher logging level so every message will
        be ignored.
        This is useful when we don't want to print anything when the logger
        is not configured by the main application.
        """
        return self._get_or_create_logger(name, level, logging.NullHandler())

    def enable_logging(self, use_json, level, log_format=DEFAULT_MESSAGE_FORMAT, stream=None):
        """ called from main after parsing arguments """
        if self._is_enabled:
            raise ValueError('re-enabling logging. Not sure whether that is ok...')

        log_level = LOG_LEVELS[level]
        logging.basicConfig(level=log_level, format=log_format, stream=stream)
        self._is_enabled = True

        self._use_json = use_json

        sys.excepthook = self._log_except_hook

        # since there could be loggers already created we go through all of them
        # and set their formatters to our custom Json formatter
        # also set their levels so they respect what the main module wants logged
        for name in self._all_names:
            logger = self.get_or_create_silent_logger(name)

            if self._use_json:
                self._make_json(logger)

        # print the start of the logging message list
        if self._use_json:
            print('[')

    def end_logging(self):
        """
        Must be called at the end of the main function
        if the caller wants json-compatible ouptut
        """
        if not self._is_enabled:
            return
        self._is_enabled = False

        # end logging
        self._all_names = set()
        logging.shutdown()

        # end json list
        if self._use_json:
            print(']')
        self._use_json = False

    def _log_except_hook(self, exctype, value, traceback):
        """
        Global hook for exceptions so we can always end logging
        """
        self.end_logging()
        sys.__excepthook__(exctype, value, traceback)

    def _get_or_create_logger(self, name, level, handler=None):
        """
        If a logger doesn't exist, we create it and set the handler,
        if it given. This avoids messing with existing loggers.
        If we are using json then we also skip adding a handler,
        since it would be removed afterwards
        """

        # logging.getLogger creates a logger if it doesn't exist,
        # so we need to check before calling it
        if handler and not self._use_json and not self._log_exists(name):
            logger = logging.getLogger(name)
            logger.addHandler(handler)
        else:
            logger = logging.getLogger(name)

        self._set_logger_level(logger, level)
        self._all_names.add(name)

        if self._use_json:
            self._make_json(logger)

        return logger

    def _make_json(self, logger):
        """
        Replace handlers of every logger by a handler that uses the JSON formatter
        """
        for handler in logger.handlers:
            logger.removeHandler(handler)
        new_handler = logging.StreamHandler(sys.stdout)
        new_handler.setFormatter(self._json_formatter)
        logger.addHandler(new_handler)

        # Don't let it propagate to parent loggers, otherwise we might get
        # duplicated messages from the root logger
        logger.propagate = False

    @staticmethod
    def _set_logger_level(logger, level):
        """
        If the logging is already initialized, we use the same level that
        was set to the root logger. This prevents imported modules' loggers
        from messing with the main module logging.
        """
        if _root_logger_wrapper.is_logging_initialized():
            logger.setLevel(_root_logger_wrapper.get_root_logger_level())
        else:
            logger.setLevel(level)

    @staticmethod
    def _log_exists(name):
        """
        We check the log manager instead of our global _all_names variable
        since the logger could have been created outside of the helper
        """
        return name in logging.Logger.manager.loggerDict
