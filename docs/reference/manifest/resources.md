# <a name="resources-element"></a><span data-ttu-id="8dee0-101">Élément Resources</span><span class="sxs-lookup"><span data-stu-id="8dee0-101">Resources element</span></span>

<span data-ttu-id="8dee0-p101">Contient des icônes, des chaînes et des URL pour le nœud [VersionOverrides](versionoverrides.md). Un élément de manifeste indique une ressource à l’aide de l’**id** de la ressource. Cela permet de conserver une taille de manifeste raisonnable, surtout lorsque les ressources sont disponibles en plusieurs versions selon les paramètres régionaux. Un **id** doit être unique au sein du manifeste et doit comporter 32 caractères au maximum.</span><span class="sxs-lookup"><span data-stu-id="8dee0-p101">Contains icons, strings, and URLs for the [VersionOverrides](versionoverrides.md) node. A manifest element specifies a resource by using the **id** of the resource. This helps to keep the size of the manifest manageable, especially when resources have versions for different locales. An **id** must be unique within the manifest and can have a maximum of 32 characters.</span></span>

<span data-ttu-id="8dee0-106">Chaque ressource peut avoir plusieurs éléments enfants **Override** afin que vous puissiez définir une ressource différente pour un paramètre régional spécifique.</span><span class="sxs-lookup"><span data-stu-id="8dee0-106">Each resource can have one or more **Override** child elements to define a different resource for a specific locale.</span></span>

## <a name="child-elements"></a><span data-ttu-id="8dee0-107">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="8dee0-107">Child elements</span></span>

|  <span data-ttu-id="8dee0-108">Élément</span><span class="sxs-lookup"><span data-stu-id="8dee0-108">Element</span></span> |  <span data-ttu-id="8dee0-109">Type</span><span class="sxs-lookup"><span data-stu-id="8dee0-109">Type</span></span>  |  <span data-ttu-id="8dee0-110">Description</span><span class="sxs-lookup"><span data-stu-id="8dee0-110">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="8dee0-111">Images</span><span class="sxs-lookup"><span data-stu-id="8dee0-111">Images</span></span>](#images)            |  <span data-ttu-id="8dee0-112">image</span><span class="sxs-lookup"><span data-stu-id="8dee0-112">image</span></span>   |  <span data-ttu-id="8dee0-113">Fournit l’URL HTTPS de l’image d’une icône.</span><span class="sxs-lookup"><span data-stu-id="8dee0-113">Provides the HTTPS URL to an image for an icon.</span></span> |
|  <span data-ttu-id="8dee0-114">**Urls**</span><span class="sxs-lookup"><span data-stu-id="8dee0-114">**Urls**</span></span>                |  <span data-ttu-id="8dee0-115">url</span><span class="sxs-lookup"><span data-stu-id="8dee0-115">url</span></span>     |  <span data-ttu-id="8dee0-p102">Fournit un URL HTTPS d’emplacement. Une URL peut comporter jusqu’à 2 048 caractères.</span><span class="sxs-lookup"><span data-stu-id="8dee0-p102">Provides an HTTPS URL location. A URL can have a maximum of 2048 characters.</span></span> |
|  <span data-ttu-id="8dee0-118">**ShortStrings**</span><span class="sxs-lookup"><span data-stu-id="8dee0-118">**ShortStrings**</span></span> |  <span data-ttu-id="8dee0-119">string</span><span class="sxs-lookup"><span data-stu-id="8dee0-119">string</span></span>  |  <span data-ttu-id="8dee0-p103">Texte pour les éléments **Label** et **Title**. Chaque élément **String** comporte 125 caractères au maximum.</span><span class="sxs-lookup"><span data-stu-id="8dee0-p103">The text for **Label** and **Title** elements. Each **String** contains a maximum of 125 characters.</span></span>|
|  <span data-ttu-id="8dee0-122">**LongStrings**</span><span class="sxs-lookup"><span data-stu-id="8dee0-122">**LongStrings**</span></span>  |  <span data-ttu-id="8dee0-123">string</span><span class="sxs-lookup"><span data-stu-id="8dee0-123">string</span></span>  | <span data-ttu-id="8dee0-p104">Texte pour les attributs **Description**. Chaque **String** comporte 250 caractères au maximum.</span><span class="sxs-lookup"><span data-stu-id="8dee0-p104">The text for **Description** attributes. Each **String** contains a maximum of 250 characters.</span></span>|

> [!NOTE]
> <span data-ttu-id="8dee0-126">Vous devez utiliser le protocole SSL (Secure Sockets Layer) pour toutes les URL dans les éléments **Image** et **Url**.</span><span class="sxs-lookup"><span data-stu-id="8dee0-126">You must use Secure Sockets Layer (SSL) for all URLs in the **Image** and **Url** elements.</span></span>

### <a name="images"></a><span data-ttu-id="8dee0-127">Images</span><span class="sxs-lookup"><span data-stu-id="8dee0-127">Images</span></span>
<span data-ttu-id="8dee0-128">Chaque icône doit disposer de trois éléments **Images**, un pour chacune des trois tailles obligatoires :</span><span class="sxs-lookup"><span data-stu-id="8dee0-128">Each icon must have three  **Images** elements, one for each of the three mandatory sizes:</span></span>

- <span data-ttu-id="8dee0-129">16x16</span><span class="sxs-lookup"><span data-stu-id="8dee0-129">16x16</span></span>
- <span data-ttu-id="8dee0-130">32x32</span><span class="sxs-lookup"><span data-stu-id="8dee0-130">32x32</span></span>
- <span data-ttu-id="8dee0-131">80x80</span><span class="sxs-lookup"><span data-stu-id="8dee0-131">80x80</span></span>

<span data-ttu-id="8dee0-132">Les tailles supplémentaires suivantes sont également prises en charge, mais ne sont pas obligatoires :</span><span class="sxs-lookup"><span data-stu-id="8dee0-132">The following additional sizes are also supported, but not required:</span></span>

- <span data-ttu-id="8dee0-133">20x20</span><span class="sxs-lookup"><span data-stu-id="8dee0-133">20x20</span></span>
- <span data-ttu-id="8dee0-134">24x24</span><span class="sxs-lookup"><span data-stu-id="8dee0-134">24x24</span></span>
- <span data-ttu-id="8dee0-135">40x40</span><span class="sxs-lookup"><span data-stu-id="8dee0-135">40x40</span></span>
- <span data-ttu-id="8dee0-136">48x48</span><span class="sxs-lookup"><span data-stu-id="8dee0-136">48x48</span></span>
- <span data-ttu-id="8dee0-137">64x64</span><span class="sxs-lookup"><span data-stu-id="8dee0-137">64x64</span></span>

> [!IMPORTANT] 
> <span data-ttu-id="8dee0-138">Outlook doit pouvoir mettre en cache les ressources d’image pour des raisons de performances.</span><span class="sxs-lookup"><span data-stu-id="8dee0-138">Important:  Outlook requires the ability to cache image resources for performance purposes.</span></span> <span data-ttu-id="8dee0-139">Par conséquent, le serveur qui héberge une ressource d’image ne doit pas ajouter les directives CACHE-CONTROL à l’en-tête de réponse.</span><span class="sxs-lookup"><span data-stu-id="8dee0-139">For this reason, the server hosting an image resource must not add any CACHE-CONTROL directives to the response header.</span></span> <span data-ttu-id="8dee0-140">Outlook la remplacerait alors automatiquement par une image générique ou par défaut.</span><span class="sxs-lookup"><span data-stu-id="8dee0-140">This will result in Outlook automatically substituting a generic or default image.</span></span>    

## <a name="resources-examples"></a><span data-ttu-id="8dee0-141">Exemples de ressources</span><span class="sxs-lookup"><span data-stu-id="8dee0-141">Resources examples</span></span> 

```XML
<Resources>
      <bt:Images>
        <bt:Image id="icon1_16x16" DefaultValue="https://www.contoso.com/icon_default.png">
          <bt:Override Locale="ja-jp" Value="https://www.contoso.com/ja-jp16-icon_default.png" />
        </bt:Image>
        <bt:Image id="icon1_32x32" DefaultValue="https://www.contoso.com/icon_default.png">
          <bt:Override Locale="ja-jp" Value="https://www.contoso.com/ja-jp32-icon_default.png" />
        </bt:Image>
        <bt:Image id="icon1_80x80" DefaultValue="https://www.contoso.com/icon_default.png">
          <bt:Override Locale="ja-jp" Value="https://www.contoso.com/ja-jp80-icon_default.png" />
        </bt:Image>
      </bt:Images>
      <bt:Urls>
        <bt:Url id="residDesktopFuncUrl" DefaultValue="https://www.contoso.com/Pages/Home.aspx">
          <bt:Override Locale="ja-jp" Value="https://www.contoso.com/Pages/Home.aspx" />
        </bt:Url>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="residLabel" DefaultValue="GetData">
          <bt:Override Locale="ja-jp" Value="JA-JP-GetData" />
        </bt:String>
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="residToolTip" DefaultValue="Get data for your document.">
          <bt:Override Locale="ja-jp" Value="JA-JP - Get data for your document." />
        </bt:String>
      </bt:LongStrings>
    </Resources>
```

```xml
<Resources>
  <bt:Images>
    <!-- Blue icon -->
    <bt:Image id="blue-icon-16" DefaultValue="YOUR_WEB_SERVER/blue-16.png"/>
    <bt:Image id="blue-icon-32" DefaultValue="YOUR_WEB_SERVER//blue-32.png"/>
    <bt:Image id="blue-icon-80" DefaultValue="YOUR_WEB_SERVER/blue-80.png"/>
  </bt:Images>
  <bt:Urls>
    <bt:Url id="functionFile" DefaultValue="YOUR_WEB_SERVER/FunctionFile/Functions.html"/>
    <!-- other URLs -->
  </bt:Urls>
  <bt:ShortStrings>
    <bt:String id="groupLabel" DefaultValue="Add-in Demo">
      <bt:Override Locale="ar-sa" Value="<Localized text>" />
    </bt:String>
    <!-- Other short strings -->
  </bt:ShortStrings>
  <bt:LongStrings>
    <bt:String id="funcReadSuperTipDescription" DefaultValue="Gets the subject of the message or appointment.">
      <bt:Override Locale="ar-sa" Value="<Localized text>." />
    </bt:String>
    <!-- Other long strings -->
  </bt:LongStrings>
</Resources>
```
