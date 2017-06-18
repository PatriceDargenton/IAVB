Intelligence Artificielle Vraiment Basique (IAVB)

# Version 3.12 du 18/06/2017
- IAVB : Intelligence Artificielle en Visual Basic -> Vraiment Basique (pour tenir compte des versions en C#) ;
- Synthèse vocale en français de Windows 10 (sous Windows 7 on n'a que l'anglais), en plus de MS-Agent ;
- Version VB.Net : mode automatique pour tester automatiquement tous les exemples, zone de liste pour la synthèse vocale ;
- Gestion des majuscules : ajout d'un booléen (bTraiterEnMinuscules) pour pouvoir comparer les résultats avec la version d'origine (si on active les accents, par exemple pour que la synthèse vocale fonctionne bien, alors il faut désactiver les majuscules, car les accents ne sont pas faciles à trouver au clavier sur les majuscules) ;
- Gestion des accents : ajout d'un booléen (bTraiterSansAccents) pour pouvoir comparer les résultats avec la version d'origine ;
- Normalisation des réponses : ajout d'un booléen (bNormalisationSortieTrimEtVbLf) pour rendre homogène les espaces et sauts de ligne dans les réponses, on avait parfois saut de ligne en trop, ou espace en trop, parfois il en manquait ;
- Version en VBA Word (code commun avec la version VB6), car il est devenu difficile de coder en VB6, l'IDE de VB6 ne s'installant pas sous Windows 10 ;
- Version VBA d'origine : fonctionnement identique à la version de 1984 (reprise des exemples pour pouvoir vérifier qu'on obtient bien les mêmes résultats) ;
- Versions VBA : correction de la copie dans le presse-papier sous Windows 64 bit (l'ancienne méthode ne fonctionnait que sous Windows 32 bits) ;
- Versions en C# : WinForm, WPF et CSharpHtml (tout juste fonctionnel, mais à terminer, il manque notamment ListBox.ScrollIntoView) ;
- Tests unitaires pour vérifier les exemples fournis ;
- Respect des règles d'analyse du code selon Visual Studio 2017 ;
- Mise au norme DotNet du code hérité de VB6 (utilisation de Generic.List(Of String), du coup on part de l'index 0 et non plus 1, ...) ;
- Version sans Microsoft.VisualBasic.dll (pour pouvoir convertir plus facilement en C#) ;
- Passage en VB 2013 (les versions 2015 et 2017 ont toujours un [bug](https://github.com/dotnet/roslyn/issues/2509) d'indentation, si on a utilisé l'indentation en mode bloc ou bien désactivé).

[IAVB.html](http://patrice.dargenton.free.fr/ia/iavb/IAVB.html) : Documentation complète
