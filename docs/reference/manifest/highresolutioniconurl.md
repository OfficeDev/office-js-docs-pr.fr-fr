# <a name="highresolutioniconurl-element"></a><span data-ttu-id="43fc6-101">Élément HighResolutionIconUrl</span><span class="sxs-lookup"><span data-stu-id="43fc6-101">HighResolutionIconUrl element</span></span>

<span data-ttu-id="43fc6-102">Spécifie l’URL de l’image qui est utilisée pour représenter votre complément Office dans l’UX d’insertion UX et l’Office Store sur les écrans à haute résolution (DPI).</span><span class="sxs-lookup"><span data-stu-id="43fc6-102">Specifies the URL of the image that is used to represent your Office Add-in in the insertion UX and Office Store on high DPI screens.</span></span>

<span data-ttu-id="43fc6-103">**Type de complément :** contenu, volet Office, courrier</span><span class="sxs-lookup"><span data-stu-id="43fc6-103">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="43fc6-104">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="43fc6-104">Syntax</span></span>

```XML
<HighResolutionIconUrl DefaultValue="string" />
```

## <a name="can-contain"></a><span data-ttu-id="43fc6-105">Peut contenir</span><span class="sxs-lookup"><span data-stu-id="43fc6-105">Can contain:</span></span>

[<span data-ttu-id="43fc6-106">Remplacement</span><span class="sxs-lookup"><span data-stu-id="43fc6-106">Override</span></span>](override.md)

## <a name="attributes"></a><span data-ttu-id="43fc6-107">Attributs</span><span class="sxs-lookup"><span data-stu-id="43fc6-107">Attributes</span></span>

|<span data-ttu-id="43fc6-108">**Attribut**</span><span class="sxs-lookup"><span data-stu-id="43fc6-108">**Attribute**</span></span>|<span data-ttu-id="43fc6-109">**Type**</span><span class="sxs-lookup"><span data-stu-id="43fc6-109">**Type**</span></span>|<span data-ttu-id="43fc6-110">**Obligatoire**</span><span class="sxs-lookup"><span data-stu-id="43fc6-110">**Required**</span></span>|<span data-ttu-id="43fc6-111">**Description**</span><span class="sxs-lookup"><span data-stu-id="43fc6-111">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="43fc6-112">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="43fc6-112">DefaultValue</span></span>|<span data-ttu-id="43fc6-113">string (URL)</span><span class="sxs-lookup"><span data-stu-id="43fc6-113">string (URL)</span></span>|<span data-ttu-id="43fc6-114">obligatoire</span><span class="sxs-lookup"><span data-stu-id="43fc6-114">required</span></span>|<span data-ttu-id="43fc6-115">Spécifie la valeur par défaut de ce paramètre, exprimée pour les paramètres régionaux spécifiés dans l’élément [DefaultLocale](defaultlocale.md).</span><span class="sxs-lookup"><span data-stu-id="43fc6-115">Specifies the default value for this setting, expressed for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|

## <a name="remarks"></a><span data-ttu-id="43fc6-116">Remarques</span><span class="sxs-lookup"><span data-stu-id="43fc6-116">Remarks</span></span>

<span data-ttu-id="43fc6-p101">Pour un complément de messagerie, l’icône apparaît dans l’interface utilisateur, sous **Fichier**  >  **Gérer les compléments**. Pour un complément de contenu ou de volet Office, l’icône apparaît dans l’interface utilisateur, sous **Insérer**  >  **Compléments**.</span><span class="sxs-lookup"><span data-stu-id="43fc6-p101">For a mail add-in, the icon is displayed in the  **File** > **Manage add-ins** UI . For a content or task pane add-in, the icon is displayed in the **Insert** > **Add-ins** UI.</span></span>

<span data-ttu-id="43fc6-119">L’image doit être dans l’un des formats de fichier suivants, avec une résolution recommandée de 64 x 64 pixels : GIF, JPG, PNG, EXIF, BMP ou TIFF.</span><span class="sxs-lookup"><span data-stu-id="43fc6-119">The image must be in one of the following file formats at a recommended resolution of 64 x 64 pixels: GIF, JPG, PNG, EXIF, BMP or TIFF.</span></span> <span data-ttu-id="43fc6-120">Pour plus d’informations, voir la section _Créer une identité visuelle cohérente pour votre application_ dans [Créer des référencements efficaces dans AppSource et dans Office](https://docs.microsoft.com/office/dev/store/create-effective-office-store-listings).</span><span class="sxs-lookup"><span data-stu-id="43fc6-120">For more information, see the section  Create a consistent visual identity for your app in Create effective Office Store apps and add-ins.</span></span>
