# <a name="sourcelocation-element"></a><span data-ttu-id="cc764-101">Élément SourceLocation</span><span class="sxs-lookup"><span data-stu-id="cc764-101">SourceLocation element</span></span>

<span data-ttu-id="cc764-p101">Spécifie les emplacements des fichiers source pour votre extension Office sous forme d’URL comprenant entre 1 et 2 018 caractères. L’emplacement source doit être une adresse HTTPS, et non un chemin d’accès de fichier.</span><span class="sxs-lookup"><span data-stu-id="cc764-p101">Specifies the source file location(s) for your Office Add-in as a URL between 1 and 2018 characters long. The source location must be an HTTPS address, not a file path.</span></span>

<span data-ttu-id="cc764-104">**Type d'extension :** contenu, volet Office, courrier</span><span class="sxs-lookup"><span data-stu-id="cc764-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="cc764-105">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="cc764-105">Syntax</span></span>

```XML
<SourceLocation DefaultValue="string" />
```

## <a name="contained-in"></a><span data-ttu-id="cc764-106">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="cc764-106">Contained in:</span></span>

- <span data-ttu-id="cc764-107">[DefaultSettings](defaultsettings.md) (compléments de contenu et extensions du volet Office)</span><span class="sxs-lookup"><span data-stu-id="cc764-107">[DefaultSettings](defaultsettings.md) (Content and task pane add-ins)</span></span>
- <span data-ttu-id="cc764-108">[FormSettings](formsettings.md) (extensions pour courrier)</span><span class="sxs-lookup"><span data-stu-id="cc764-108">[FormSettings](formsettings.md) (Mail add-ins)</span></span>
- <span data-ttu-id="cc764-109">[ExtensionPoint](extensionpoint.md) (extensions pour courriers contextuels)</span><span class="sxs-lookup"><span data-stu-id="cc764-109">[ExtensionPoint](extensionpoint.md) (Contextual mail add-ins)</span></span>

## <a name="can-contain"></a><span data-ttu-id="cc764-110">Peut contenir</span><span class="sxs-lookup"><span data-stu-id="cc764-110">Can contain:</span></span>

[<span data-ttu-id="cc764-111">remplacement</span><span class="sxs-lookup"><span data-stu-id="cc764-111">Override</span></span>](override.md)

## <a name="attributes"></a><span data-ttu-id="cc764-112">Attributs</span><span class="sxs-lookup"><span data-stu-id="cc764-112">Attributes</span></span>

|<span data-ttu-id="cc764-113">**Attribut**</span><span class="sxs-lookup"><span data-stu-id="cc764-113">**Attribute**</span></span>|<span data-ttu-id="cc764-114">**Type**</span><span class="sxs-lookup"><span data-stu-id="cc764-114">**Type**</span></span>|<span data-ttu-id="cc764-115">**requis**</span><span class="sxs-lookup"><span data-stu-id="cc764-115">**Required**</span></span>|<span data-ttu-id="cc764-116">**Description**</span><span class="sxs-lookup"><span data-stu-id="cc764-116">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="cc764-117">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="cc764-117">DefaultValue</span></span>|<span data-ttu-id="cc764-118">URL</span><span class="sxs-lookup"><span data-stu-id="cc764-118">URL</span></span>|<span data-ttu-id="cc764-119">requis</span><span class="sxs-lookup"><span data-stu-id="cc764-119">required</span></span>|<span data-ttu-id="cc764-120">Spécifie la valeur par défaut de ce paramètre pour les paramètres régionaux spécifiés dans l’élément [DefaultLocale](defaultlocale.md).</span><span class="sxs-lookup"><span data-stu-id="cc764-120">Specifies the default value for this setting for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|
