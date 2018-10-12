# <a name="supporturl-element"></a><span data-ttu-id="8aa02-101">Élément SupportUrl</span><span class="sxs-lookup"><span data-stu-id="8aa02-101">SupportUrl element</span></span>

<span data-ttu-id="8aa02-102">Spécifie l’URL d’une page qui fournit des informations de prise en charge pour votre complément.</span><span class="sxs-lookup"><span data-stu-id="8aa02-102">Specifies the URL of a page that provides support information for your add-in.</span></span>

## <a name="syntax"></a><span data-ttu-id="8aa02-103">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="8aa02-103">Syntax</span></span>

```XML
<OfficeApp>
...
  <IconUrl DefaultValue="https://contoso.com/assets/icon-32.png" />
  <HighResolutionIconUrl DefaultValue="https://contoso.com/assets/hi-res-icon.png"/>
  
  
  <SupportUrl DefaultValue="https://contoso.com/support " />
  
  
  <AppDomains>
  ...
  </AppDomains>
...
</OfficeApp>
```

## <a name="contained-in"></a><span data-ttu-id="8aa02-104">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="8aa02-104">Contained in:</span></span>

[<span data-ttu-id="8aa02-105">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="8aa02-105">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="8aa02-106">Peut contenir</span><span class="sxs-lookup"><span data-stu-id="8aa02-106">Can contain:</span></span>

|  <span data-ttu-id="8aa02-107">Élément</span><span class="sxs-lookup"><span data-stu-id="8aa02-107">Element</span></span> | <span data-ttu-id="8aa02-108">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="8aa02-108">Required</span></span> | <span data-ttu-id="8aa02-109">Description</span><span class="sxs-lookup"><span data-stu-id="8aa02-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="8aa02-110">Remplacement</span><span class="sxs-lookup"><span data-stu-id="8aa02-110">Override</span></span>](override.md)   | <span data-ttu-id="8aa02-111">Non</span><span class="sxs-lookup"><span data-stu-id="8aa02-111">No</span></span> | <span data-ttu-id="8aa02-112">Spécifie le paramètre pour les URL de paramètres régionaux supplémentaires</span><span class="sxs-lookup"><span data-stu-id="8aa02-112">Specifies the setting for additional locale urls</span></span> |

## <a name="attributes"></a><span data-ttu-id="8aa02-113">Attributs</span><span class="sxs-lookup"><span data-stu-id="8aa02-113">Attributes</span></span>

|<span data-ttu-id="8aa02-114">**Attribut**</span><span class="sxs-lookup"><span data-stu-id="8aa02-114">**Attribute**</span></span>|<span data-ttu-id="8aa02-115">**Type**</span><span class="sxs-lookup"><span data-stu-id="8aa02-115">**Type**</span></span>|<span data-ttu-id="8aa02-116">**Obligatoire**</span><span class="sxs-lookup"><span data-stu-id="8aa02-116">**Required**</span></span>|<span data-ttu-id="8aa02-117">**Description**</span><span class="sxs-lookup"><span data-stu-id="8aa02-117">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="8aa02-118">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="8aa02-118">DefaultValue</span></span>|<span data-ttu-id="8aa02-119">URL</span><span class="sxs-lookup"><span data-stu-id="8aa02-119">URL</span></span>|<span data-ttu-id="8aa02-120">obligatoire</span><span class="sxs-lookup"><span data-stu-id="8aa02-120">required</span></span>|<span data-ttu-id="8aa02-121">Spécifie la valeur par défaut de ce paramètre, exprimée pour les paramètres régionaux spécifiés dans l’élément [DefaultLocale](defaultlocale.md).</span><span class="sxs-lookup"><span data-stu-id="8aa02-121">Specifies the default value for this setting, expressed for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|
