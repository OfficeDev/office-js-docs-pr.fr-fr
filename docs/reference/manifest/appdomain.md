# <a name="appdomain-element"></a><span data-ttu-id="201ae-101">AppDomain, élément</span><span class="sxs-lookup"><span data-stu-id="201ae-101">AppDomain element</span></span>

<span data-ttu-id="201ae-102">Indique un domaine supplémentaire permettant de charger des pages dans la fenêtre du complément.</span><span class="sxs-lookup"><span data-stu-id="201ae-102">Specifies an additional domain that will be used to load pages in the add-in window.</span></span>

<span data-ttu-id="201ae-103">**Type de complément :** application de contenu, de volet Office, de messagerie</span><span class="sxs-lookup"><span data-stu-id="201ae-103">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="201ae-104">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="201ae-104">Syntax</span></span>

```XML
<AppDomain>string</AppDomain>
```

> [!IMPORTANT]
> <span data-ttu-id="201ae-105">La valeur de l’élément**AppDomain**doit inclure le protocole (par exemple,`<AppDomain>https://myappdomain<AppDomain>`).</span><span class="sxs-lookup"><span data-stu-id="201ae-105">The value of the **AppDomain** element must include the protocol (e.g., `<AppDomain>https://myappdomain<AppDomain>`).</span></span>

## <a name="contained-in"></a><span data-ttu-id="201ae-106">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="201ae-106">Contained in</span></span>

[<span data-ttu-id="201ae-107">AppDomains</span><span class="sxs-lookup"><span data-stu-id="201ae-107">AppDomains</span></span>](appdomains.md)

## <a name="remarks"></a><span data-ttu-id="201ae-108">Remarques</span><span class="sxs-lookup"><span data-stu-id="201ae-108">Remarks</span></span>

<span data-ttu-id="201ae-109">Les éléments **AppDomain** sont utilisés pour indiquer les domaines supplémentaires autres que celui spécifié dans l’[élément SourceLocation](sourcelocation.md).</span><span class="sxs-lookup"><span data-stu-id="201ae-109">The  AppDomains and **AppDomain** elements are used to specify any additional domains other than the one specified in the [SourceLocation element. For more information, see Office Add-ins XML manifest](sourcelocation.md).</span></span> <span data-ttu-id="201ae-110">Pour plus d’informations, reportez-vous au [manifeste XML de compléments Office](/office/dev/add-ins/develop/add-in-manifests).</span><span class="sxs-lookup"><span data-stu-id="201ae-110">For more information, see [Office Add-ins XML manifest](/office/dev/add-ins/develop/add-in-manifests).</span></span>
