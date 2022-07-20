---
title: Codes d'erreur de l'API commune de l'Office
description: Cet article documente les messages d’erreur que vous pouvez rencontrer lors de l’utilisation de l’API Office Common.
ms.date: 02/08/2022
ms.localizationpriority: medium
ms.openlocfilehash: 80637e1573e120664ee89685a9c9579dbcfe6b2e
ms.sourcegitcommit: df7964b6509ee6a807d754fbe895d160bc52c2d3
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/20/2022
ms.locfileid: "66889610"
---
# <a name="office-common-api-error-codes"></a>Codes d'erreur de l'API commune de l'Office

Cet article documente les messages d’erreur que vous pouvez rencontrer lors de l’utilisation du modèle d’API commune. Ces codes d’erreur ne s’appliquent pas aux API spécifiques à l’application, telles que l’API JavaScript Excel ou l’API JavaScript Word.

Consultez [les modèles d’API](../develop/understanding-the-javascript-api-for-office.md#api-models) pour en savoir plus sur les différences entre l’API commune et les modèles d’API spécifiques à l’application.

## <a name="error-codes"></a>Codes d’erreur

Le tableau suivant répertorie les codes d’erreur, les noms et les messages affichés, ainsi que les conditions qu’ils indiquent.

|**Code d’erreur**|**Nom de l’erreur**|**Message de l’erreur**|**Condition**|
|:-----|:-----|:-----|:-----|
|1000|Le type de forçage de type est incorrect|Le type de forçage de type spécifié n’est pas pris en charge.|Le type de contrainte n’est pas pris en charge dans l’application Office. (Par exemple, les types de forçage de type OOXML et  HTML ne sont pas pris en charge dans Excel.)|
|1001|Une erreur s’est produite lors de la lecture des données|La sélection actuelle n’est pas prise en charge.|La sélection actuelle de l’utilisateur n’est pas prise en charge (autrement dit, cela ne correspond pas aux types de forçage de type pris en charge).|
|1002|Le type de forçage de type est incorrect|Le type de forçage de type spécifié n’est pas compatible avec ce type de liaison.|Le développeur de solutions a fourni une combinaison incompatible de type de forçage de type et de type de liaison.|
|1003|Une erreur s’est produite lors de la lecture des données|Les valeurs rowCount ou columnCount spécifiées sont incorrectes.|L’utilisateur fournit un nombre de lignes ou de colonnes incorrect.|
|1004|Une erreur s’est produite lors de la lecture des données|La sélection actuelle n’est pas compatible avec le type de forçage de type spécifié.|La sélection actuelle n’est pas prise en charge pour le type de forçage de type spécifié par cette application.|
|1005|Une erreur s’est produite lors de la lecture des données|Les valeurs startRow ou startColumn spécifiées sont incorrectes.|L’utilisateur fournit des valeurs startRow ou startCol incorrectes.|
|1006|Une erreur s’est produite lors de la lecture des données|Les paramètres de coordonnées ne peuvent pas être utilisés avec le type de forçage de type « Tableau » lorsque le tableau contient des cellules fusionnées.|L’utilisateur essaie d’obtenir des données partielles à partir d’un tableau non uniforme (c’est-à-dire un tableau qui contient des cellules fusionnées.) |
|1007|Une erreur s’est produite lors de la lecture des données|La taille du document est trop importante.|L’utilisateur tente d’obtenir un document plus grand que la taille actuellement prise en charge.|
|1008|Une erreur s’est produite lors de la lecture des données|La taille du jeu de données demandé est trop importante.|L’utilisateur demande à lire des données au-delà des limites de données définies par l’application Office.|
|1009|Une erreur s’est produite lors de la lecture des données|Le type de fichier spécifié n’est pas pris en charge.|L’utilisateur envoie un type de fichier incorrect.|
|2000|Une erreur s’est produite lors de l’écriture des données|Le type d’objet de données fourni n’est pas pris en charge. |Un objet de données non pris en charge est fourni.|
|2001|Une erreur s’est produite lors de l’écriture des données|Impossible d’écrire dans la sélection actuelle.|La sélection actuelle de l’utilisateur n’est pas prise en charge pour une opération d’écriture. (Par exemple, lorsque l’utilisateur sélectionne une image.)|
|2002|Une erreur s’est produite lors de l’écriture des données|L’objet de données fourni n’est pas compatible avec la forme ou les dimensions de la sélection actuelle.|Plusieurs cellules sont sélectionnées (et la forme de sélection ne correspond pas à la forme des données). Plusieurs cellules sont sélectionnées (et les dimensions de la sélection ne correspondent pas aux dimensions des données).|
|2003|Une erreur s’est produite lors de l’écriture des données|L’opération SET a échoué, car l’objet de données fourni remplacera les données.|Une seule cellule est sélectionnée et l’objet de données fourni remplace les données dans la feuille de calcul.|
|2004|Une erreur s’est produite lors de l’écriture des données|L’objet de données fourni ne correspond pas à la taille de la sélection actuelle.|L’utilisateur fournit un objet plus grand que la taille de la sélection actuelle.|
|2005|Une erreur s’est produite lors de l’écriture des données|Les valeurs startRow ou startColumn spécifiées sont incorrectes.|L’utilisateur fournit des valeurs startRow ou startCol incorrectes.|
|2006|Une erreur de format incorrect s’est produite|Le format de l’objet de données spécifié est incorrect.|Le développeur de solutions fournit une chaîne HTML ou OOXML incorrecte, une chaîne HTML au format incorrect ou une chaîne  OOXML incorrecte.|
|2007|L’objet de données est incorrect|Le type de l’objet de données spécifié n’est pas compatible avec la sélection actuelle.|Le développeur de solutions fournit un objet de données qui n’est pas compatible avec le type de forçage de type spécifié.|
|2008|Une erreur s’est produite lors de l’écriture des données|TBD|TBD|
|2009|Une erreur s’est produite lors de l’écriture des données|L’objet de données spécifié est trop volumineux.|L’utilisateur tente de définir des données au-delà des limites de données définies par l’application Office.|
|2010|Une erreur s’est produite lors de l’écriture des données|Les paramètres de coordonnées ne peuvent pas être utilisés avec le type de forçage de type Tableau lorsque le tableau contient des cellules fusionnées.|L’utilisateur tente de définir des données partielles à partir d’un tableau non uniforme (c’est-à-dire un tableau qui contient des cellules fusionnées.)|
|3000|Une erreur s’est produite lors de la création de la liaison|Impossible d’effectuer de liaison avec la sélection actuelle.|La sélection de l’utilisateur n’est pas prise en charge pour la liaison. (Par exemple, l’utilisateur sélectionne une image ou un autre objet non pris en charge.)|
|3001|Une erreur s’est produite lors de la création de la liaison|TBD|TBD|
|3002|Erreur de liaison incorrecte|La liaison spécifiée n’existe pas.|Le développeur tente de créer une liaison avec une liaison non existante ou supprimée.|
|3003|Une erreur s’est produite lors de la création de la liaison|Les sélections non contiguës ne sont pas prises en charge.|L’utilisateur effectue des sélections multiples.|
|3004|Une erreur s’est produite lors de la création de la liaison|Impossible de créer une liaison avec la sélection actuelle et le type de liaison spécifié.|Il existe plusieurs conditions dans lesquelles cela pourrait se produire. Consultez la section « Conditions d’erreur de création de liaison » plus loin dans cet article.|
|3005|Opération de liaison incorrecte|Ce type de liaison ne prend pas en charge cette action.|Le développeur envoie une opération d’ajout de ligne ou d’ajout de colonne sur un type de liaison qui n’est pas de type `table`forçage.|
|3006|Une erreur s’est produite lors de la création de la liaison|L’élément nommé n’existe pas.|L’élément nommé est introuvable. Aucune table ni aucun contrôle de contenu portant ce nom n’existe.|
|3007|Une erreur s’est produite lors de la création de la liaison|Nous avons trouvé plusieurs objets du même nom.|Erreur de collision : plusieurs contrôles de contenu portant le même nom existent et l’échec lors de la collision est défini `true`sur .|
|3008|Une erreur s’est produite lors de la création de la liaison|Le type de liaison spécifié n’est pas compatible avec l’élément nommé fourni.|L’élément nommé ne peut pas être lié au type. Par exemple, un contrôle de contenu contient du texte, mais le développeur a essayé de lier à l’aide du type `table`de contrainte .|
|3009|Opération de liaison incorrecte|Le type de liaison n’est pas pris en charge.|Utilisé pour la compatibilité descendante.|
|3010|Opération de liaison non prise en charge|Le contenu sélectionné doit être dans un format de tableau. Placez les données sous forme de tableau, puis réessayez.|Le développeur tente d’utiliser la ou `deleteAllDataValuesAsync` les `addRowsAsync` méthodes de l’objet `TableBinding` sur des données de type `matrix`forçage.|
|4000|Une erreur s’est produite lors de la lecture des paramètres|Le nom de paramètre spécifié n’existe pas.|Un nom de paramètre non existant est fourni.|
|4001|Une erreur s’est produite lors de l’enregistrement des paramètres|Les paramètres n’ont pas pu être enregistrés.|Les paramètres n’ont pas pu être enregistrés.|
|4002|Une erreur relative à des paramètres périmés s’est produite|Les paramètres n’ont pas pu être enregistrés car ils sont périmés.|Les paramètres sont périmés et le développeur a indiqué de ne pas les remplacer.|
|5000|Une erreur relative à des paramètres périmés s’est produite|L’opération n’est pas prise en charge.|L’opération n’est pas prise en charge dans l’application Office actuelle. Par exemple, `document.getSelectionAsync` est appelé à partir d’Outlook.|
|5001|Erreur interne|Une erreur interne s’est produite.|Fait référence à une condition d’erreur interne qui peut survenir pour l’une des raisons suivantes.<br/><table><tr><td>Un complément utilisé par un autre utilisateur partageant le classeur a créé une liaison quasiment au même moment et votre complément doit recommencer le processus de liaison.</tr></td><tr><td>Une erreur inconnue s’est produite.</tr></td><tr><td>L’opération a échoué.</tr></td><tr><td>L’accès a été refusé car l’utilisateur n’est pas membre d’un rôle autorisé.</tr></td><tr><td>L’accès a été refusé car une communication chiffrée sécurisée est exigée.</tr></td><tr><td>Les données sont obsolètes et l’utilisateur doit confirmer l’activation des requêtes pour les actualiser.</tr></td><tr><td>Le quota d’UC de la collection de sites est dépassé.</tr></td><tr><td>Le quota de mémoire de la collection de sites est dépassé.</tr></td><tr><td>Le quota de mémoire de la session est dépassé.</tr></td><tr><td>Le classeur est dans un état non valide et l’opération ne peut pas être effectuée.</tr></td><tr><td>La session a expiré car elle était inactive et l’utilisateur doit recharger le classeur.</tr></td><tr><td>Le nombre maximal de sessions autorisées par utilisateur est dépassé.</tr></td><tr><td>L’opération a été annulée par l’utilisateur.</tr></td><tr><td>L’opération ne peut pas aboutir car elle prend trop de temps.</tr></td><tr><td>La demande ne peut pas aboutir et une nouvelle tentative doit être effectuée.</tr></td><tr><td>La période d’évaluation du produit a expiré.</tr></td><tr><td>La session a expiré car elle était inactive.</tr></td><tr><td>L’utilisateur n’est pas autorisé à effectuer l’opération sur la plage spécifiée.</tr></td><tr><td>Les paramètres régionaux de l’utilisateur ne correspondent pas à la session de collaboration active.</tr></td><tr><td>L’utilisateur n’est plus connecté et doit actualiser ou rouvrir le classeur.</tr></td><tr><td>La plage demandée n’existe pas dans la feuille.</tr></td><tr><td>L’utilisateur n’est pas autorisé à modifier le classeur.</tr></td><tr><td>Le classeur ne peut pas être modifié car il est verrouillé.</tr></td><tr><td>La session ne peut pas enregistrer automatiquement le classeur.</tr></td><tr><td>La session ne peut pas actualiser son verrouillage du fichier du classeur.</tr></td><tr><td>La demande ne peut pas être traitée et une nouvelle tentative doit être effectuée.</tr></td><tr><td>Les informations de connexion de l’utilisateur n’ont pas pu être vérifiées et doivent être saisies de nouveau.</tr></td><tr><td>L’accès a été refusé à l’utilisateur.</tr></td><tr><td>Le classeur partagé doit être mis à jour.</tr></td></table>|
|5002|Autorisation refusée|L’opération demandée n’est pas autorisée sur le mode de document actuel.|Le développeur de solutions soumet une opération de définition, mais le document est dans un mode qui n’autorise pas de modifications, telles que « Restreindre la modification ».|
|5003|Une erreur s’est produite lors de l’enregistrement de l’événement|Le type d’événement spécifié n’est pas pris en charge par l’objet actuel.|Le développeur de solutions tente d’inscrire ou d’annuler l’inscription d’un gestionnaire pour un événement qui n’existe pas.|
|5004|L’appel d’API est incorrect|L’appel d’API n’est pas correct dans le contexte actuel.|Un appel non valide est effectué pour le contexte, par exemple, en essayant d’utiliser un `CustomXMLPart` objet dans Excel.|
|5005|Données périmées|Échec de l’opération car les données sur le serveur sont périmées.|Les données sur le serveur doivent être actualisées.|
|5006|Expiration de la session|La session de document a expiré. Rechargez le document. |La session a expiré.|
|5007|L’appel d’API est incorrect|L’énumération n’est pas prise en charge dans le contexte actuel.|L’énumération n’est pas prise en charge dans le contexte actuel.|
|5009|Autorisation refusée|Accès refusé|Le complément n’est pas autorisé à appeler l’API spécifique.|
|5012|Session non valide ou expirée|Votre session de navigateur Office a expiré ou n’est pas valide. Pour continuer, actualisez la page.|La session entre le client Office et le serveur a expiré, ou la date, l’heure ou le fuseau horaire est incorrect sur votre ordinateur.|
|6000|Nœud incorrect|Le nœud spécifié est introuvable.|Le `CustomXmlPart` nœud est introuvable.|
|6100|Une erreur relative à du code XML personnalisé s’est produite.|Une erreur relative à du code XML personnalisé s’est produite.|L’appel d’API est incorrect.|
|7000|ID incorrect|L’ID spécifié n’existe pas.|L’ID est incorrect.|
|7001|Navigation non valide|L’objet se trouve à un emplacement dans lequel la navigation n’est pas prise en charge.|L’utilisateur peut trouver l’objet, mais ne peut pas naviguer jusqu’à celui-ci. (Par exemple, dans Word, la liaison est effectuée avec l’en-tête, le pied de page ou un commentaire.)|
|7002|Navigation non valide|L’objet est verrouillé ou protégé.|L’utilisateur tente d’accéder à une plage verrouillée ou protégée.|
|7004|Navigation non valide|Échec de l’opération car l’index est hors limites.|L’utilisateur tente d’accéder à un index qui est hors limites.|
|8000|Paramètre manquant|Nous n’avons pas pu mettre en forme la cellule de tableau, car certaines valeurs de paramètre sont manquantes. Vérifiez à nouveau les paramètres et réessayez.|Il manque certains paramètres à la méthode cellFormat. Par exemple, il manque les paramètres cells, format ou tableOptions.|
|8010|Valeur non valide|Un ou plusieurs paramètres cells contiennent des valeurs qui ne sont pas autorisées. Vérifiez les valeurs et réessayez.|L’énumération de référence des cellules communes n’est pas définie. Par exemple, Tout, Données, En-têtes.|
|8011|Valeur non valide|Un ou plusieurs paramètres tableOptions contiennent des valeurs qui ne sont pas autorisées. Vérifiez les valeurs et réessayez.|Une des valeurs saisies dans tableOptions n’est pas valide.|
|8012|Valeur non valide|Un ou plusieurs paramètres format contiennent des valeurs qui ne sont pas autorisées. Vérifiez les valeurs et réessayez.|Une des valeurs de format n’est pas valide.|
|8020|En dehors de la plage|La valeur d’index de ligne se trouve en dehors de la plage autorisée. Utilisez une valeur (supérieure ou égale à 0) inférieure au nombre de lignes.|L’index de ligne est supérieur à l’index de ligne le plus élevé du tableau ou est inférieur à 0.|
|8021|En dehors de la plage|La valeur d’index de colonne se trouve en dehors de la plage autorisée. Utilisez une valeur (supérieure ou égale à 0) inférieure au nombre de colonnes.|L’index de colonne est supérieur à l’index de colonne le plus élevé du tableau ou est inférieur à 0.|
|8022|En dehors de la plage|La valeur se trouve en dehors de la plage autorisée.|Certaines des valeurs dans le format se trouvent en dehors des plages prises en charge.|
|9016|Autorisation refusée|Autorisation refusée|L’accès est refusé.|
|9020|Erreur de réponse générique|Une erreur interne s’est produite.|Fait référence à une condition d’erreur interne, qui peut se produire pour plusieurs raisons.|
|9021|Erreur d’enregistrement|Une erreur de connexion s’est produite lors de la tentative d’enregistrement de l’élément sur le serveur.|Impossible d’enregistrer l’élément. Cela peut être dû à une erreur de connexion au serveur si vous utilisez le mode en ligne dans le bureau Outlook, ou à une tentative de réenregistrage d’un brouillon d’élément qui a été supprimé du serveur Exchange.|
|9022|Message d’erreur dans un autre magasin|Impossible de récupérer l’ID EWS, car le message est enregistré dans un autre magasin.|L’ID EWS du message actuel n’a pas pu être récupéré, car le message a peut-être été déplacé ou la boîte aux lettres d’envoi a peut-être changé.|
|9041|Erreur réseau|L’utilisateur n’est plus connecté au réseau. Veuillez vérifier la connexion réseau et réessayer.|L’utilisateur n’a plus accès au réseau ou à Internet.|
|9043|Type de pièce jointe non pris en charge|Le type de pièce jointe n’est pas pris en charge.|L’API ne prend pas en charge le type de pièce jointe. Par exemple, `item.getAttachmentContentAsync` génère cette erreur si la pièce jointe est une image incorporée au format texte enrichi ou s’il s’agit d’un type d’élément autre qu’un e-mail ou un élément de calendrier (par exemple, un contact ou un élément de tâche).|
|12002|||Un des éléments suivants :<br> - Aucune page n’existe à l’URL qui a été transmise à `displayDialogAsync`.<br> - La page qui a été transmise à `displayDialogAsync` a été chargée, mais la boîte de dialogue a été redirigée vers une page introuvable ou impossible à charger, ou a été redirigée vers une URL dont la syntaxe n’est pas valide. Erreur signalée dans la boîte de dialogue. Elle déclenche un événement `DialogEventReceived` dans la page hôte.|
|12003|||La boîte de dialogue a été redirigée vers une URL avec le protocole HTTP. C’est le protocole HTTPS qui est requis. Erreur signalée dans la boîte de dialogue. Elle déclenche un événement `DialogEventReceived` dans la page hôte.|
|12004|||Le domaine de l’URL transmis à `displayDialogAsync` n’est pas approuvé. Le domaine doit être le même domaine que celui de la page hôte (y compris le protocole et le numéro de port). Erreur générée par l’appel de `displayDialogAsync`.|
|12005|||L’URL transmise à `displayDialogAsync` utilise le protocole HTTP. C’est le protocole HTTPS qui est requis. Erreur générée par l’appel de `displayDialogAsync`. (Dans certaines versions d’Office, le message d’erreur renvoyé avec le code 12005 est identique à celui renvoyé avec le code 12004.)|
|12006|||La boîte de dialogue a été fermée, généralement parce que l’utilisateur choisit le bouton **X**. Erreur signalée dans la boîte de dialogue. Elle déclenche un événement `DialogEventReceived` dans la page hôte.|
|12007|||Une boîte de dialogue est déjà ouverte à partir de cette fenêtre hôte. Une fenêtre hôte, par exemple un volet Office, ne peut avoir qu’une seule boîte de dialogue ouverte à la fois. Erreur générée par l’appel de `displayDialogAsync`.|
|12009|||L’utilisateur a choisi d’ignorer la boîte de dialogue. Cette erreur peut se produire dans les versions en ligne d’Office, quand les utilisateurs peuvent choisir d’autoriser ou non un complément à afficher une boîte de dialogue. Erreur générée par l’appel de `displayDialogAsync`.|
|13nnn|||Consultez [Causes et gestion des erreurs de getAccessToken](../develop/troubleshoot-sso-in-office-add-ins.md#causes-and-handling-of-errors-from-getaccesstoken).|

## <a name="binding-creation-error-conditions"></a>Conditions d’erreur de création de liaison

Lorsqu’une liaison est créée dans l’API, indiquez le type de liaison que vous souhaitez utiliser. Le tableau suivant répertorie les types de liaisons et les comportements de liaison attendus.

### <a name="behavior-in-excel"></a>Comportement dans Excel

Le tableau suivant résume le comportement de liaison dans Excel.

|**Type de liaison spécifié**|**Sélection réelle**|**Comportement**|
|:-----|:-----|:-----|
|Matrice|Plage de cellules (y compris dans un tableau et une cellule unique)|Une liaison de type `matrix` est créée sur les cellules sélectionnées. Aucune modification dans le document n’est attendue.|
|Matrice|Texte sélectionné dans la cellule|Une liaison de type `matrix` est créée sur l’ensemble de la cellule. Aucune modification dans le document n’est attendue.|
|Matrice|Sélection multiple/sélection incorrecte (par exemple, l’utilisateur sélectionne une image, un objet ou un objet Word Art.)|Impossible de créer la liaison.|
|Tableau|Plage de cellules (y compris une cellule unique)|Impossible de créer la liaison.|
|Tableau|Plage de cellules dans un tableau (comprend une seule cellule dans un tableau, le tableau entier, ou du texte dans la cellule d’un tableau)|Une liaison est créée dans le tableau entier.|
|Tableau|Demi-sélection dans un tableau et demie sélection en dehors du tableau|Impossible de créer la liaison.|
|Tableau|Texte sélectionné dans la cellule (pas dans le tableau)|Impossible de créer la liaison.|
|Tableau|Sélection multiple/sélection incorrecte (par exemple, l’utilisateur sélectionne une image, un objet, un objet Word Art, etc.)|Impossible de créer la liaison.|
|Texte|Plage de cellules|Impossible de créer la liaison.|
|Texte|Plage de cellules dans un tableau|Impossible de créer la liaison.|
|Texte|Cellule unique|Une liaison de type `text` est créée.|
|Texte|Cellule unique dans un tableau|Une liaison de type `text` est créée.|
|Texte|Texte sélectionné dans la cellule|Une liaison de type `text` dans la cellule entière est créée.|

### <a name="behavior-in-word"></a>Comportement dans Word

Le tableau suivant résume le comportement de liaison dans Word.

|**Type de liaison spécifié**|**Sélection réelle**|**Comportement**|
|:-----|:-----|:-----|
|Matrice|Texte|Impossible de créer la liaison.|
|Matrice|Tableau entier|Une liaison de type `matrix` est créée. Le document est modifié et un contrôle de contenu doit encapsuler la table. |
|Matrice|Plage dans un tableau|Impossible de créer la liaison.|
|Matrice|Sélection non valide (par exemple, objets multiples, incorrects, etc.)|Impossible de créer la liaison.|
|Tableau|Texte|Impossible de créer la liaison.|
|Tableau|Tableau entier|Une liaison de type `text` est créée.|
|Tableau|Plage dans un tableau|Impossible de créer la liaison.|
|Tableau|Sélection non valide (par exemple, objets multiples, incorrects, etc.)|Impossible de créer la liaison.|
|Texte|Tableau entier|Une liaison de type `text` est créée.|
|Texte|Plage dans un tableau|Impossible de créer la liaison.|
|Texte|Sélection multiple|La dernière sélection sera encapsulée avec un contrôle de contenu et une liaison à ce contrôle. Un contrôle de contenu de type `text` est créé.|
|Texte|Sélection non valide (par exemple, objets multiples, incorrects, etc.)|Impossible de créer la liaison.|

## <a name="see-also"></a>Voir aussi

- [Cycle de vie du développement des compléments Office](../overview/office-add-ins.md)
- [Compréhension de l’API JavaScript pour Office](../develop/understanding-the-javascript-api-for-office.md)
- [Gestion des erreurs avec les API JavaScript spécifiques à l’application](../testing/application-specific-api-error-handling.md)
- [Résolution des problèmes de messages d’erreur pour l’authentification unique (SSO)](../develop/troubleshoot-sso-in-office-add-ins.md)
- [Résolution des erreurs rencontrées par l’utilisateur avec des compléments Office](../testing/troubleshoot-development-errors.md)
