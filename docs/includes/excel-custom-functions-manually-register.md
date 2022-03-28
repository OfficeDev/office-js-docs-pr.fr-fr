Si l’espace de noms `CONTOSO` n’est pas disponible dans le menu de saisie semi-automatique, procédez comme suit pour inscrire le complément dans Excel.

### <a name="excel-on-windows-or-mac"></a>[Excel sur Windows ou Mac](#tab/excel-windows)

1. Dans Excel, choisissez l’onglet **Insérer** , puis la flèche vers le bas située à droite de **Mes compléments**.

    :::image type="content" source="../images/select-insert.png" alt-text="Capture d’écran du ruban Insérer dans Excel sur Windows, avec la flèche vers le bas mes compléments mise en surbrillance.":::

1. Dans la liste des compléments disponibles, recherchez la section **Compléments de développeur**, puis sélectionnez le complément **starcount** pour effectuer cette opération.

    :::image type="content" source="../images/list-starcount.png" alt-text="Capture d’écran du ruban Insérer dans Excel sous Windows, avec le complément Fonctions personnalisées Excel mis en évidence dans la liste Mes compléments.":::

# <a name="excel-on-the-web"></a>[Excel sur le web](#tab/excel-online)

1. Dans Excel, sélectionnez l’onglet **Insertion**, puis **Compléments**.

    :::image type="content" source="../images/excel-cf-online-register-add-in-1.png" alt-text="Capture d’écran du ruban Insertion dans Excel sur le web, avec le bouton Mes compléments mise en évidence.":::

1. Sélectionnez **Gérer mes Compléments** et sélectionnez **Télécharger mon complément**.

1. Sélectionnez **Parcourir...** et accédez au répertoire racine du projet créé par le Générateur de Yo Office.

1. Sélectionnez le fichier **manifest.xml** puis sélectionnez **Ouvrir**, puis sélectionnez **Télécharger**.

1. Essayez la nouvelle fonction. Dans la cellule **B1**, tapez le texte **=CONTOSO. GETSTARCOUNT(« OfficeDev », « Excel-Custom-Functions »)**, puis appuyez sur Entrée. Le résultat dans la cellule **B1** doit correspondre au nombre d’étoiles actuellement attribuées au [référentiel GitHub Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions).

---
