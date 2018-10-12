# <a name="requestedheight-element"></a><span data-ttu-id="27951-101">Élément RequestedHeight</span><span class="sxs-lookup"><span data-stu-id="27951-101">RequestedHeight element</span></span>

<span data-ttu-id="27951-102">Spécifie la hauteur initiale (en pixels) d’un complément de contenu ou le complément messagerie.</span><span class="sxs-lookup"><span data-stu-id="27951-102">Specifies the initial height (in pixels) of a content add-in or mail add-in.</span></span> 

<span data-ttu-id="27951-103">**Type de complément :** Contenu, messagerie</span><span class="sxs-lookup"><span data-stu-id="27951-103">**Add-in type:** Content, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="27951-104">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="27951-104">Syntax</span></span>

```XML
<RequestedHeight>integer</RequestedHeight>
```

## <a name="contained-in"></a><span data-ttu-id="27951-105">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="27951-105">Contained in:</span></span>

- <span data-ttu-id="27951-106">[DefaultSettings](defaultsettings.md) (Contenu compléments) avec une valeur qui peut être comprise entre 32 et 1 000</span><span class="sxs-lookup"><span data-stu-id="27951-106">[DefaultSettings](defaultsettings.md) (Content add-ins) with a value that can be between 32 and 1000</span></span>
- <span data-ttu-id="27951-107">[DesktopSettings](desktopsettings.md) et [TabletSettings](tabletsettings.md) (Compléments de messagerie) avec une valeur qui peut être comprise entre 32 et 450</span><span class="sxs-lookup"><span data-stu-id="27951-107">[DesktopSettings](desktopsettings.md) and [TabletSettings](tabletsettings.md) (Mail add-ins) with a value that can be between 32 and 450</span></span>
- <span data-ttu-id="27951-108">[ExtensionPoint](extensionpoint.md) (Compléments de messagerie contextuelle) avec une valeur contenue entre 140 et 450 pour le point d'extension **DetectedEntity** et entre 32 et 450 pour le point d'extension **CustomPane**</span><span class="sxs-lookup"><span data-stu-id="27951-108">[ExtensionPoint](extensionpoint.md) (Contextual mail add-ins) with a value that can be between 140 and 450 for the **DetectedEntity** extension point and between 32 and 450 for the **CustomPane** extension point</span></span>