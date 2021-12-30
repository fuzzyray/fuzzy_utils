import json
from pathlib import Path


class PersistentDict(dict):
    """A simple implementation of a Python dictionary with a persistent backing store"""

    def __init__(self, *args, backing_store=None, replace=False, **kwargs):
        self._backing_store = backing_store
        if backing_store:
            if Path(backing_store).is_file():
                if not replace:
                    with open(backing_store, "r", encoding="utf-8") as f:
                        try:
                            super().__init__(**json.load(f))
                        except TypeError:
                            raise TypeError(
                                f"backing_store: '{backing_store}' does not contain a well-formed dictionary object"
                            ) from None
                else:
                    super().__init__(*args, **kwargs)
                    with open(backing_store, "w", encoding="utf-8") as f:
                        json.dump(self, f)
            elif not Path(backing_store).exists():
                super().__init__(*args, **kwargs)
                with open(backing_store, "w", encoding="utf-8") as f:
                    json.dump(self, f)
            else:
                raise TypeError(f"backing_store: '{backing_store}' exists, but is not a file")
        else:
            super().__init__(*args, **kwargs)

    def __delitem__(self, key):
        super().__delitem__(key)
        if self._backing_store:
            with open(self.backing_store, "w", encoding="utf-8") as f:
                json.dump(self, f)

    def __setitem__(self, key, value):
        super().__setitem__(key, value)
        if self._backing_store:
            with open(self.backing_store, "w", encoding="utf-8") as f:
                json.dump(self, f)

    @property
    def backing_store(self):
        return self._backing_store
