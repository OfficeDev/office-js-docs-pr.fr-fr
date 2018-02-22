---
title: Test de l’intégration de Script Lab
description: ''
ms.date: 12/04/2017
---


# <a name="testing-script-lab-integration"></a>Test de l’intégration de Script Lab

Il s’agit d’un exemple de fichier de test destiné à décrire une fonctionnalité ScriptLab à venir qui permettra aux développeurs de tester leurs extraits de code dans Excel, Word ou PowerPoint.  

## <a name="prerequisites"></a>Conditions préalables
- Vous aurez besoin d’une URL d’affichage issue d’un extrait de code Script Lab.
- Remarque : nous *devons* indiquer que Script Lab a besoin d’Office 365 pour explorer les extraits de code les plus récents. Les développeurs peuvent obtenir un abonnement à Office 365 via notre [programme pour les développeurs Office 365](https://dev.office.com/devprogram), uniquement à des fins de développement.  


## <a name="try-it-out-button"></a>Bouton « Essayez ! »
De cette façon, nous ajoutons un bouton « Essayez ! ». Nous vous recommandons d’y associer un extrait de code.  Pour ce faire, nous utilisons une classe de la structure d’interface utilisateur d’Office permettant de définir un lien sous forme de bouton. Sur le lien lui-même, n’oubliez pas de définir l’attribut *aria label*.

**Démonstration :**

<a href="https://bornholm-edge.azurewebsites.net/#/view/gist/excel/0cc24cee687141d1c2726c0feea70911" class="ms-Button" aria-label="Open this snippet in Script Lab, an Office Add-in">Essayez !</a>


<button href="https://bornholm-edge.azurewebsites.net/#/view/gist/excel/0cc24cee687141d1c2726c0feea70911" class="ms-Button" aria-label="Open this snippet in Script Lab, an Office Add-in">Essayez !</button>


**Code :**
```html
<a href="ahttps://bornholm-edge.azurewebsites.net/#/view/gist/excel/0cc24cee687141d1c2726c0feea70911" class="ms-Button" aria-label="Open this snippet in Script Lab, an Office Add-in">Try it out</a>
```



## <a name="embed-script-lab-as-an-iframe"></a>Incorporer un laboratoire de scripts sous la forme d’un IFrame
Dans ce mode, nous incorporons un extrait de code directement sous la forme d’un IFrame dans nos documents. La largeur a été définie sur 95 % (par rapport à la largeur de tous les autres extraits) et nous vous recommandons de supprimer la bordure de cadre de l’IFrame.  En général, la hauteur doit être ajustée pour correspondre à l’extrait de code.

**Démonstration :**
<iframe src="https://bornholm-edge.azurewebsites.net/#/view/gist/excel/0cc24cee687141d1c2726c0feea70911" height="600px" width="95%" frameborder="0"></iframe>

**Code :**
```html
<iframe src="https://bornholm-edge.azurewebsites.net/#/view/gist/excel/0cc24cee687141d1c2726c0feea70911" height="600px" width="95%" frameborder="0"></iframe>
```

## <a name="testing-considerations"></a>Considérations relatives aux tests
Nous devons vérifier les abonnements mobiles autres qu’Office 365 (nous avons reçu des commentaires sur les documents office.js indiquant que de nombreux développeurs utilisaient la version 2013 ou une version antérieure).  

Pour le chemin d’accès de l’incorporation, une déconnexion finale est nécessaire. Nous devons également nous assurer que le contenu figurant dans la page d’affichage de liste répond à nos instructions d’accessibilité.

## <a name="see-also"></a>Voir aussi
