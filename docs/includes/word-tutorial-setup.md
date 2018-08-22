Vous commencerez ce didacticiel par la configuration de votre projet de développement. 

> [!NOTE]
> Cette page décrit une étape individuelle d’un didacticiel sur les compléments Word. Si vous êtes arrivé à cette page via les résultats du moteur de recherche ou d’un autre lien direct, accédez à la page d’introduction du [didacticiel sur les compléments Word](../tutorials/word-tutorial.yml) pour démarrer le didacticiel à partir du début.

> [!TIP]
> Si vous ne l’avez pas déjà fait, lisez l’article [Créer votre premier complément Word](../quickstarts/word-quickstart.md?tabs=visual-studio-code). Vous devez notamment savoir comment charger une version test d’un complément Word.

## <a name="prerequisites"></a>Conditions préalables

Pour utiliser ce didacticiel, les logiciels suivants doivent être installés. 

- Word 2016, version 1711 (Démarrer en un clic version 8730.1000) ou version ultérieure. Vous devrez peut-être participer au programme Office Insider pour obtenir cette version. Pour plus d’informations, reportez-vous à [Participez au programme Office Insider](https://products.office.com/office-insider?tab=tab-1).
- [Nœud et npm](https://nodejs.org/en/) 
- [Git Bash](https://git-scm.com/downloads) (ou un autre client Git)

## <a name="setup"></a>Installation

1. Clonez le référentiel GitHub du [didacticiel sur les compléments Word](https://github.com/OfficeDev/Word-Add-in-Tutorial).
2. Ouvrez une fenêtre Git Bash, ou une invite système Node.JS, et accédez au dossier **Start** du projet.
3. Exécutez la commande `npm install` pour installer les outils et les bibliothèques répertoriées dans le fichier package.json. 
4. Effectuez les étapes décrites dans la rubrique relative à l’[ajout de certificats auto-signés comme certificat racine approuvé](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) pour approuver le certificat pour le système d’exploitation de votre ordinateur de développement.

