## VBA-challege
# Module 2 - VBA challenge
During the challenge I referenced the following to aid in the understanding and completion of the assignment:
* Reviewed all the class activites in GitLab Working Folder - reperforming a majority of them
* Watched the following YouTube videos from Dr. A in our #03-resources Slack channel:
  - [VBA Bonus Demo 01 - Using Sheet references as Variables in VBA Scripting](https://www.youtube.com/watch?v=SIGr245Yb6M)
  - [VBA Bonus Demo 02 - Doing calculations in multiple sheets (Updated)](https://www.youtube.com/watch?v=mear99YPXSA)
  - [VBA Bonus Demo 03 - Aggregates (Part 1) - Uses the attached 'armada' excel workbook](https://www.youtube.com/watch?v=NcnCckeLNao)
  - [VBA Scripting Unit Day 03 Activity #7 - US Census (Part 1)](https://www.youtube.com/watch?v=IO1MSYqeBfo)
  - [VBA Scripting Unit Day 03 Activity #8 - US Census (Part 2)](https://www.youtube.com/watch?v=tBAlDe0oei0)
* I used google to search for VBA references related to:
  - [For Each](https://excelchamps.com/vba/loop-sheets/)
  - [MATCH()](https://excelchamps.com/vba/match/)
* I also used google to search for assistance on how to understand the yearly price change calculation in VBA, particularly around the opening price:
  - [Sub WorksheetsLoop() by ibaloyan](https://github.com/ibaloyan/Stock_Analysis_with_VBA)
    *   ' declare opening price as Double
    *     Dim openingPrice As Double
    *     openingPrice = 0

    *   ' retrieve next ticker symbol's opening price
    *     openingPrice = ws.Cells(Row + 1, 3).Value
