# 40Kalculator

This is a probability calculator for the game Warhammer 40,000 (40k)

****Purpose****

Many actions in 40k involve complex probability calcualtions. It can be possible to estimate the expected value, but this is not as useful as a full probability distribution, since the expected value typically only has a probability of around 50%. When the result of the game depends on it, you might instead want to know the result with greater certainty. This can be obtained from a complete probability distribution.

This application will take the probability values for a particular attack against a particular target and produce a probability distribution of the possible damage that could result from that attack.

****Features****

- Generates a probability distribution for an attack against a single target, or against multiple models with a wounds characteristic from 1 to 6
- Accounts for the following rules:
  - Random attacks
  - Random damage
  - Lethal Hits
  - Sustained Hits (fixed amount only)
  - Devastating Wounds
  - Modified AP on critical wounds
      - Variable crit chance accounted for
  - Feel no Pain saves
- Rollover wounds correctly respecting devastating wounds and Feel no Pain saves
- Graphical display of Probability Distribution
- DataGridView displaying complete probability distribution. Compatible to copy and past to Excel.
- Probabilities calculated for a specific result and cumulative ("at least this good")
- Selecting results on the graphical display or DataGridView highlights the corresponding result in the other display

****Usage****

- Each calculation is for a specific attack profile against a specific target.
- Probabilities can be entered as:
  - Decimal
  - Fractions (using / )
  - Dice up/down notation (ie. 3+, 5- representing 3 or better and 5 or worse)
    - A dropdown box allows you to select different types of dice (D3 to D20). This will only affect the Dice up/down notation.
   
****Language and building****

The project is coded in VB.net using Visual Studio. The repository should contain everything necessary to build the project. You should clone the repository using the built in tool in Visual Studio.

These are the recommended settings to build the .exe

Configuration: Release | Any CPU

Target framework: net6.0-windows

Deployment mode: Self-contained

Target runtime: win-x86

Options: Produce single file

