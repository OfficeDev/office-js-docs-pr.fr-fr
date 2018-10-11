# <a name="allowsnapshot-element"></a><span data-ttu-id="496ae-101">Élément AllowSnapshot</span><span class="sxs-lookup"><span data-stu-id="496ae-101">AllowSnapshot element</span></span>

<span data-ttu-id="496ae-102">Indique si une capture instantanée de votre complément de contenu est enregistrée avec le document hôte.</span><span class="sxs-lookup"><span data-stu-id="496ae-102">Specifies whether a snapshot image of your content add-in is saved with the host document.</span></span>

<span data-ttu-id="496ae-103">**Type de complément :** Contenu</span><span class="sxs-lookup"><span data-stu-id="496ae-103">**Add-in type:** Content</span></span>

## <a name="syntax"></a><span data-ttu-id="496ae-104">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="496ae-104">Syntax</span></span>

```XML
<AllowSnapshot> [true | false]</AllowSnapshot>
```

## <a name="contained-in"></a><span data-ttu-id="496ae-105">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="496ae-105">Contained in:</span></span>

[<span data-ttu-id="496ae-106">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="496ae-106">OfficeApp</span></span>](officeapp.md)

## <a name="remarks"></a><span data-ttu-id="496ae-107">Remarques</span><span class="sxs-lookup"><span data-stu-id="496ae-107">Remarks</span></span>

 > [!IMPORTANT]
 > <span data-ttu-id="496ae-108">**AllowSnapshot** est `true` par défaut.</span><span class="sxs-lookup"><span data-stu-id="496ae-108">Security Note:**AllowSnapshot** is true`true` by default.</span></span> <span data-ttu-id="496ae-109">Cela crée une image du complément visible pour les utilisateurs qui ouvrent le document dans une version de l’application hôte ne prenant pas en charge les compléments Office, ou fournissant une image statique du complément si l’application hôte ne peut pas se connecter au serveur qui héberge le complément.</span><span class="sxs-lookup"><span data-stu-id="496ae-109">This makes an image of the add-in visible for users that open the document in a version of the host application that doesn't support Office Add-ins, or provides a static image of the add-in if the host application can't connect to the server hosting the add-in.</span></span> <span data-ttu-id="496ae-110">Toutefois, cela signifie également que les informations potentiellement sensibles affichées dans le complément sont accessibles directement à partir du document hébergeant le complément.</span><span class="sxs-lookup"><span data-stu-id="496ae-110">However, this also means that potentially sensitive information displayed in the add-in can be accessed directly from the document hosting the add-in.</span></span>

