# <a name="requestedheight-element"></a>Élément RequestedHeight

Spécifie la hauteur initiale (en pixels) d’un complément de contenu ou le complément messagerie. 

**Type de complément :** Contenu, messagerie

## <a name="syntax"></a>Syntaxe

```XML
<RequestedHeight>integer</RequestedHeight>
```

## <a name="contained-in"></a>Contenu dans

- [DefaultSettings](defaultsettings.md) (Contenu compléments) avec une valeur qui peut être comprise entre 32 et 1 000
- [DesktopSettings](desktopsettings.md) et [TabletSettings](tabletsettings.md) (Compléments de messagerie) avec une valeur qui peut être comprise entre 32 et 450
- [ExtensionPoint](extensionpoint.md) (Compléments de messagerie contextuelle) avec une valeur contenue entre 140 et 450 pour le point d'extension **DetectedEntity** et entre 32 et 450 pour le point d'extension **CustomPane**