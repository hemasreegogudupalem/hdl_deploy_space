import contextlib
import io
import unittest
from unittest import mock

import run_frontend


class FrontendLauncherTests(unittest.TestCase):
    def test_reuses_an_existing_hdl_server_instead_of_binding_again(self):
        output = io.StringIO()
        with (
            mock.patch.object(run_frontend, "_is_port_in_use", return_value=True),
            mock.patch.object(
                run_frontend,
                "_is_hdl_server_running",
                return_value=True,
            ),
            mock.patch.object(run_frontend.uvicorn, "run") as run_server,
            contextlib.redirect_stdout(output),
        ):
            run_frontend.main()

        run_server.assert_not_called()
        self.assertIn("already running", output.getvalue())

    def test_reports_when_another_application_owns_the_port(self):
        with (
            mock.patch.object(run_frontend, "_is_port_in_use", return_value=True),
            mock.patch.object(
                run_frontend,
                "_is_hdl_server_running",
                return_value=False,
            ),
        ):
            with self.assertRaisesRegex(SystemExit, "another application"):
                run_frontend.main()


if __name__ == "__main__":
    unittest.main()
