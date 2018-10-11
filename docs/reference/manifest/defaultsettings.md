# <a name="defaultsettings-element"></a><span data-ttu-id="12199-101">Élément DefaultSettings</span><span class="sxs-lookup"><span data-stu-id="12199-101">DefaultSettings element</span></span>

<span data-ttu-id="12199-102">Spécifie l’emplacement source par défaut et d’autres paramètres par défaut pour votre complément de contenu ou de volet Office.</span><span class="sxs-lookup"><span data-stu-id="12199-102">Specifies the default source location and other default settings for your content or task pane add-in .</span></span>

<span data-ttu-id="12199-103">**Type de complément :** contenu, volet Office</span><span class="sxs-lookup"><span data-stu-id="12199-103">**Add-in type:** Content, Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="12199-104">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="12199-104">Syntax</span></span>

```XML
<DefaultSettings>
  ...
</DefaultSettings>
```

## <a name="contained-in"></a><span data-ttu-id="12199-105">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="12199-105">Contained in:</span></span>

[<span data-ttu-id="12199-106">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="12199-106">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="12199-107">Peut contenir</span><span class="sxs-lookup"><span data-stu-id="12199-107">Can contain:</span></span>

|<span data-ttu-id="12199-108">**Élément**</span><span class="sxs-lookup"><span data-stu-id="12199-108">**Element**</span></span>|<span data-ttu-id="12199-109">**Contenu**</span><span class="sxs-lookup"><span data-stu-id="12199-109">**Content**</span></span>|<span data-ttu-id="12199-110">**Courrier**</span><span class="sxs-lookup"><span data-stu-id="12199-110">**Mail**</span></span>|<span data-ttu-id="12199-111">**Volet Office**</span><span class="sxs-lookup"><span data-stu-id="12199-111">\*\*\*\* Taskpane</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="12199-112">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="12199-112">SourceLocation</span></span>](sourcelocation.md)|<span data-ttu-id="12199-113">x</span><span class="sxs-lookup"><span data-stu-id="12199-113">x</span></span>||<span data-ttu-id="12199-114">x</span><span class="sxs-lookup"><span data-stu-id="12199-114">x</span></span>|
|[<span data-ttu-id="12199-115">RequestedWidth</span><span class="sxs-lookup"><span data-stu-id="12199-115">RequestedWidth</span></span>](requestedwidth.md)|<span data-ttu-id="12199-116">x</span><span class="sxs-lookup"><span data-stu-id="12199-116">x</span></span>|||
|[<span data-ttu-id="12199-117">RequestedHeight</span><span class="sxs-lookup"><span data-stu-id="12199-117">RequestedHeight</span></span>](requestedheight.md)|<span data-ttu-id="12199-118">x</span><span class="sxs-lookup"><span data-stu-id="12199-118">x</span></span>|||

## <a name="remarks"></a><span data-ttu-id="12199-119">Remarques</span><span class="sxs-lookup"><span data-stu-id="12199-119">Remarks</span></span>

<span data-ttu-id="12199-120">L’emplacement source et les autres paramètres de l’élément **DefaultSettings** s’appliquent uniquement aux compléments de volet Office et de contenu. Pour les compléments de messagerie, vous spécifiez les emplacements par défaut pour les fichiers sources et d’autres paramètres par défaut dans l’élément [FormSettings](formsettings.md).</span><span class="sxs-lookup"><span data-stu-id="12199-120">The source location and other settings in the  **DefaultSettings** element apply only to content and task pane add-ins. For mail add-ins, you specify the default locations for source files and other default settings in the [FormSettings](formsettings.md) element.</span></span>

