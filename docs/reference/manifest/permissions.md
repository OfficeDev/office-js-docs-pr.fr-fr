# <a name="permissions-element"></a><span data-ttu-id="b2f5c-101">Élément Permissions</span><span class="sxs-lookup"><span data-stu-id="b2f5c-101">Permissions element</span></span>

<span data-ttu-id="b2f5c-102">Spécifie le niveau d’accès à l’API de votre complément Office ; vous devez demander des autorisations basées sur le principe des privilèges minimaux.</span><span class="sxs-lookup"><span data-stu-id="b2f5c-102">Specifies the level of API access for your Office Add-in; you should request permissions based on the principle of least privilege.</span></span>

<span data-ttu-id="b2f5c-103">**Type de complément :** contenu, volet Office, messagerie</span><span class="sxs-lookup"><span data-stu-id="b2f5c-103">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="b2f5c-104">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="b2f5c-104">Syntax</span></span>

<span data-ttu-id="b2f5c-105">Pour les compléments du volet de tâches et de contenu :</span><span class="sxs-lookup"><span data-stu-id="b2f5c-105">For content and task pane add-ins:</span></span>

```XML
 <Permissions> [Restricted | ReadDocument | ReadAllDocument | WriteDocument | ReadWriteDocument]</Permissions>
```

<span data-ttu-id="b2f5c-106">Pour les compléments de messagerie</span><span class="sxs-lookup"><span data-stu-id="b2f5c-106">For mail add-ins:</span></span>

```XML
 <Permissions>[Restricted | ReadItem | ReadWriteItem | ReadWriteMailbox]</Permissions>
```

## <a name="contained-in"></a><span data-ttu-id="b2f5c-107">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="b2f5c-107">Contained in:</span></span>

[<span data-ttu-id="b2f5c-108">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="b2f5c-108">OfficeApp</span></span>](officeapp.md)

## <a name="remarks"></a><span data-ttu-id="b2f5c-109">Remarques</span><span class="sxs-lookup"><span data-stu-id="b2f5c-109">Remarks</span></span>

<span data-ttu-id="b2f5c-110">Pour plus de détails, consultez l’article relatif à la [demande d’autorisations pour utiliser des API dans des compléments de contenu et de volet Office](https://docs.microsoft.com/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins) et celui décrivant les [autorisations de complément Outlook](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="b2f5c-110">For more detail, see [Requesting permissions for API use in content and task pane add-ins](https://docs.microsoft.com/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins) and [Understanding Outlook add-in permissions](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>
