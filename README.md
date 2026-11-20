# Tri — classification d'événements moteur par deux LLM

Un seul fichier Python, `tri_llm.py`, sans interface graphique.

Pour chaque ligne d'un classeur Excel, il lit les colonnes **D** « Description »,
**E** « Immediate Action », **F** « Long term », demande à deux modèles de la
passerelle LiteLLM interne dans quelle catégorie classer l'événement, et compare
leurs réponses :

| Colonne | Contenu |
|---|---|
| **H** « LLM Mistral » | catégorie choisie par `mistral-small-4` |
| **I** « LLM Gemma » | catégorie choisie par `gemma-4` |
| **K** « Comparaison » | `OK` blanc sur vert si H = I, `KO` blanc sur rouge sinon, `ERREUR` blanc sur orange si un appel a échoué |

Catégories : `engine shutdown commanded`, `engine power modulated`,
`engine forced to idle`, et `non classé` si aucune ne convient.

## Installation

```
pip install -r requirements.txt
copy Tri_keys.env.example Tri_keys.env
```

Puis coller dans `Tri_keys.env` les deux clés de la page « Clés API » de
LiteMaaS (une clé par modèle) : `TRI_MISTRAL_KEY` et `TRI_GEMMA_KEY`.

## Utilisation

```
python tri_llm.py                         -> demande le nom du fichier
python tri_llm.py mon_fichier.xlsx        -> fichier donné
python tri_llm.py mon_fichier.xlsx --limit 20    (essai sur 20 lignes)
python tri_llm.py mon_fichier.xlsx --force       (tout recommencer)
```

Le résultat est écrit dans `<nom>_LLM.xlsx` à côté de l'original, qui n'est
jamais modifié. Un journal `<nom>_LLM.log` est écrit au même endroit.

- **Reprise** : si `<nom>_LLM.xlsx` existe déjà, les lignes classées sont
  conservées et seules les lignes manquantes ou en `ERREUR` sont renvoyées aux
  modèles. Sauvegarde toutes les 10 lignes ; `Ctrl+C` sauvegarde et s'arrête.
- **Débit** : les clés LiteMaaS sont limitées (20 requêtes/minute par défaut).
  Le script espace ses appels en conséquence et réessaie sur passerelle
  saturée. Comptez environ 3 secondes par ligne. Réglable : `TRI_MISTRAL_RPM`,
  `TRI_GEMMA_RPM`.
- **Lignes vides** (D, E et F vides) : ignorées.
- **TLS** : si `CERTIFICATE_VERIFY_FAILED`, renseigner `DALLEM_CA_BUNDLE`
  (ou `DALLEM_VERIFY_SSL=false`) dans `Tri_keys.env`.

Toutes les options sont décrites en tête de `tri_llm.py` et dans
`Tri_keys.env.example`.
