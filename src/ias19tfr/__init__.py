"""Motore di valutazione IAS 19 del TFR."""
from .engine import proietta, valuta
from .model import Dipendente, Ipotesi, IpotesiFinanziarie, Risultato, TavolaPermanenza

__all__ = ["proietta", "valuta", "Dipendente", "Ipotesi", "IpotesiFinanziarie", "Risultato", "TavolaPermanenza"]
__version__ = "0.1.0"
