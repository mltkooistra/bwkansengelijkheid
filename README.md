# BW & kansengelijkheid – Spoor 1

**Vraag:** Wat is de relatie tussen tevredenheid met het leven, perspectief op de toekomst en gevoel van invloed/grip?

## Draaien

Open een Anaconda-prompt in deze map en voer uit:

```
python spoor1_analyse.py
```

Het script leest standaard `O:\Projecten lopend\BW & kansengelijkheid\Data\Werkbestand BW en kansengelijkheid.sav`.
Wil je een ander bestand gebruiken, geef dan het pad mee: `python spoor1_analyse.py "pad\naar\bestand.sav"`.

Daarna:
1. Lees eerst **`spoor1_checks.txt`**. Daarin staan de kolommen die gevonden zijn, hoe antwoorden zijn gehercodeerd en de frequenties. Controleer vooral de stellingen, het netwerk ('weet ik niet' telt als 'nee') en de vergelijking met `Q_levenalgemeen`.
2. **`spoor1_factsheet_data.xlsx`** gaat naar de vormgever. Tabblad `0_Inhoud` noemt per tabblad de subvraag en een voorstel voor de visual, plus de definities.

## Tabbladen per subvraag

| Subvraag | Tabbladen |
|---|---|
| 1. Hangt (on)tevredenheid samen met de toekomstblik (+ grip)? Correlatie en regressie | 1a–1f (incl. 1f per leeftijdsgroep voor 'jongeren') |
| 2. Welke groepen (bijv. ontevreden & pessimistisch)? Waarover zijn ze pessimistisch? Wie zitten erin? | 2a–2g (2e = achtergrond, 2f = verleden/jeugd) |
| 3. Grotere kans op pessimisme als de basisdingen niet op orde zijn? | 3a, 3b |
| 4. Wat zijn de basisdingen? | 4_Basisdingen |
| 5. Is zekerheid nu iets anders dan zekerheid over de toekomst? | 5a–5c |
| 6. Invloed / grip (agency) | 6a invloed per domein · 6b 'vastzitten' (onvoldoende én geen invloed) · 6c buffert grip/netwerk tegen pessimisme? · 6d regressie toekomstverwachting |
| Extra verhalen | 7a uitkomsten naar achtergrond en jeugd ('kansenkaart') · 7b/7c redenen achter pessimisme/optimisme (gezondheid, wonen, sociaal) |

`spoor1_open_antwoorden_INTERN.xlsx` bevat de open antwoorden ('andere zorgen voor de toekomst') per groep, bedoeld om citaten uit te kiezen. **Niet zo doorsturen**: check eerst of er herleidbare informatie in staat.

## Definities (aan te passen bovenin het script)

- Tevreden = cijfer ≥ 7 voor 'leven in het algemeen' (`TEVREDEN_GRENS`)
- Domein niet op orde = cijfer < 6 (`ONVOLDOENDE_GRENS`)
- Pessimistisch = verwacht dat het over 5 jaar (een beetje of veel) slechter gaat
- Alles is gewogen met `weegfactor` (provincie × leeftijd × opleiding, zie `weegfactoren.xlsx`). n = ongewogen.

## Overig

- `_archief/` bevat de oude notebook, de eenmalige fix-/testscripts en de oude figuren. Die zijn niet meer nodig en staan ook in git. De map mag weg.
