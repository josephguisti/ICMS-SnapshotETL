"""
icms_etl_core_vectorized.py
~~~~~~~~~~~~~~~~~~~~~~~~~~~
Historical A/B reference module.

The vectorized master rebuild now lives in ``icms_etl_core``. This module
re-exports the production functions so existing diagnostics can continue to
compare paths without duplicating implementation logic.
"""

from __future__ import annotations

from icms_etl_core import (
    _build_master_from_histories as _build_master_from_histories_vectorized,
    rebuild_client_project_master,
    record_master_inputs,
    update_client_project_master,
)

__all__ = [
    "_build_master_from_histories_vectorized",
    "record_master_inputs",
    "rebuild_client_project_master",
    "update_client_project_master",
]
