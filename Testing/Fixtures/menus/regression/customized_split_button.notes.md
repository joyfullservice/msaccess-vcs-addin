# customized_split_button

Regression fixture for a built-in split-button popup (Type 13, Id 212) with a
customized submenu: hidden built-in children, a custom blank button, and a
duplicate built-in child Id.

Export must keep `BuiltIn: true` and `Id: 212` (not downgrade to an Id-less
replica). Import recreates via `Add(13, 212)` and injects serialized children
into the empty fresh submenu.
