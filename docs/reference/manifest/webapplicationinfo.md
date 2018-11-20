# <a name="webapplicationinfo-element"></a><span data-ttu-id="2df9e-101">Élément WebApplicationInfo</span><span class="sxs-lookup"><span data-stu-id="2df9e-101">WebApplicationInfo element</span></span>

<span data-ttu-id="2df9e-102">Prend en charge l’authentification unique (SSO) dans des compléments Office. Cet élément contient des informations sur le complément sous deux formes :</span><span class="sxs-lookup"><span data-stu-id="2df9e-102">Supports single sign-on (SSO) in Office Add-ins. This element contains information for the add-in as both:</span></span>

- <span data-ttu-id="2df9e-103">En tant que *ressource* OAuth 2.0 pour laquelle l’application Office peut requérir des autorisations.</span><span class="sxs-lookup"><span data-stu-id="2df9e-103">An OAuth 2.0 *resource* to which the Office host application might need permissions.</span></span>
- <span data-ttu-id="2df9e-104">Un *client* OAuth 2.0 pouvant requérir des autorisations dans Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="2df9e-104">An OAuth 2.0 *client* that might need permissions to Microsoft Graph.</span></span>

> [!NOTE]
> <span data-ttu-id="2df9e-105">L’API Authentification unique est actuellement prise en charge en préversion pour Word, Excel, Outlook et PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="2df9e-105">The Single Sign-on API is currently supported in preview for Word, Excel, Outlook, and PowerPoint.</span></span> <span data-ttu-id="2df9e-106">Pour en savoir plus sur les plateformes qui prennent en charge l’API Authentification unique, consultez l’article� [Ensembles de conditions requises de l’API d’identité](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets?view=office-js).</span><span class="sxs-lookup"><span data-stu-id="2df9e-106">For more information about where the Single Sign-on API is currently supported, see [IdentityAPI requirement sets](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets?view=office-js).</span></span> <span data-ttu-id="2df9e-107">Si vous utilisez un complément Outlook, veillez à activer l’authentification moderne pour la location d’Office 365.</span><span class="sxs-lookup"><span data-stu-id="2df9e-107">If you are working with an Outlook add-in, be sure to enable Modern Authentication for the Office 365 tenancy.</span></span> <span data-ttu-id="2df9e-108">Pour savoir comment procéder, consultez l’article relatif à� l’[activation du client pour l’authentification moderne dans Exchange Online](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).</span><span class="sxs-lookup"><span data-stu-id="2df9e-108">For information about how to do this, see [Exchange Online: How to enable your tenant for modern authentication](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).</span></span>

<span data-ttu-id="2df9e-109">**WebApplicationInfo** est un élément enfant de l’élément [VersionOverrides](versionoverrides.md) dans le manifeste.</span><span class="sxs-lookup"><span data-stu-id="2df9e-109">**WebApplicationInfo** is a child element of the [VersionOverrides](versionoverrides.md) element in the manifest.</span></span>  

## <a name="child-elements"></a><span data-ttu-id="2df9e-110">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="2df9e-110">Child elements</span></span>

|  <span data-ttu-id="2df9e-111">Élément</span><span class="sxs-lookup"><span data-stu-id="2df9e-111">Element</span></span> |  <span data-ttu-id="2df9e-112">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="2df9e-112">Required</span></span>  |  <span data-ttu-id="2df9e-113">Description</span><span class="sxs-lookup"><span data-stu-id="2df9e-113">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="2df9e-114">**Id**</span><span class="sxs-lookup"><span data-stu-id="2df9e-114">**Id**</span></span>    |  <span data-ttu-id="2df9e-115">Oui</span><span class="sxs-lookup"><span data-stu-id="2df9e-115">Yes</span></span>   |  <span data-ttu-id="2df9e-116">**ID d’application** du service associé au complément, tel qu’inscrit dans le point de terminaison Azure Active Directory (Azure AD) v2.0.</span><span class="sxs-lookup"><span data-stu-id="2df9e-116">The **Application Id** of the add-in's associated service as registered in the Azure Active Directory v 2.0 endpoint.</span></span>|
|  <span data-ttu-id="2df9e-117">**Resource**</span><span class="sxs-lookup"><span data-stu-id="2df9e-117">**Resource**</span></span>  |  <span data-ttu-id="2df9e-118">Oui</span><span class="sxs-lookup"><span data-stu-id="2df9e-118">Yes</span></span>   |  <span data-ttu-id="2df9e-119">Spécifie l’**URI de l’ID d’application** du complément, tel qu’inscrit dans le point de terminaison Azure AD v2.0.</span><span class="sxs-lookup"><span data-stu-id="2df9e-119">Specifies the **Application ID URI** of the add-in as registered in the Azure Active Directory v 2.0 endpoint.</span></span>|
|  [<span data-ttu-id="2df9e-120">Scopes</span><span class="sxs-lookup"><span data-stu-id="2df9e-120">Scopes</span></span>](scopes.md)                |  <span data-ttu-id="2df9e-121">Non</span><span class="sxs-lookup"><span data-stu-id="2df9e-121">No</span></span>  |  <span data-ttu-id="2df9e-122">Spécifie les autorisations dont le complément a besoin pour Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="2df9e-122">Specifies the permissions that the add-in needs to Microsoft Graph.</span></span>  |

> [!NOTE] 
> <span data-ttu-id="2df9e-123">À l’heure actuelle, il est nécessaire que les ressources de votre complément correspondent à son hôte.</span><span class="sxs-lookup"><span data-stu-id="2df9e-123">Note: Currently, it's necessary that your add-in's Resource matches its Host.</span></span> <span data-ttu-id="2df9e-124">Office ne demandera pas un jeton pour un complément à moins de pouvoir prouver qu’il en est le propriétaire ; à l’heure actuelle, ceci s’effectue en hébergeant le complément sous le nom de domaine complet de la ressource.</span><span class="sxs-lookup"><span data-stu-id="2df9e-124">Office will not request a Token for an add-in unless it can prove ownership, and today this is done by hosting the add-in under the Resource's fully-qualified domain name.</span></span>

## <a name="webapplicationinfo-example"></a><span data-ttu-id="2df9e-125">Exemple pour WebApplicationInfo</span><span class="sxs-lookup"><span data-stu-id="2df9e-125">WebApplicationInfo example</span></span>

```xml
<OfficeApp>
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    ...
    <WebApplicationInfo>
      <Id>12345678-abcd-1234-efab-123456789abc</Id>
      <Resource>api://myDomain.com/12345678-abcd-1234-efab-123456789abc<Resource>
      <Scopes>
        <Scope>Files.Read.All</Scope>
        <Scope>offline_access</Scope>
        <Scope>openid</Scope>
        <Scope>profile</Scope>        
      </Scopes>
    </WebApplicationInfo>
  </VersionOverrides>
...
</OfficeApp>
```
