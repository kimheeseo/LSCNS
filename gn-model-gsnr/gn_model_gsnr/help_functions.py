"""Unit-conversion helpers used by the GN/GSNR model."""

from __future__ import annotations

import numpy as np

ArrayLike = float | int | np.ndarray


def dB2Lin(value: ArrayLike) -> np.ndarray:
    """Convert a power ratio from dB to linear units."""
    return np.power(10.0, np.asarray(value, dtype=float) / 10.0)


def Lin2dB(value: ArrayLike) -> np.ndarray:
    """Convert a positive linear power ratio to dB."""
    arr = np.asarray(value, dtype=float)
    if np.any(arr <= 0):
        raise ValueError("Linear values must be strictly positive for dB conversion.")
    return 10.0 * np.log10(arr)


def dBm2Lin(value_dbm: ArrayLike) -> np.ndarray:
    """Convert dBm to watts."""
    return np.power(10.0, (np.asarray(value_dbm, dtype=float) - 30.0) / 10.0)


def Lin2dBm(value_w: ArrayLike) -> np.ndarray:
    """Convert positive watts to dBm."""
    arr = np.asarray(value_w, dtype=float)
    if np.any(arr <= 0):
        raise ValueError("Power in watts must be strictly positive for dBm conversion.")
    return 10.0 * np.log10(arr) + 30.0
