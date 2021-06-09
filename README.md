# SE_lottery
This is the script, that makes Student Equimpment lottery at UNIS happening.

## Requirements for running this script
This script needs a few python packages to work properly. Namely `pandas`, `xlwt` and `openpyxl`. You can install them by hand, but I recommend using a anaconda environment. Anaconda can be downloaded [here](https://www.anaconda.com/products/individual). You can create a new environment calles `SE_lottery` by running the command
```
conda create -n SE_lottery python=3.6 pandas xlwt openpyxl
```

## How to run a lottery
1. Download the answers from the excel form from sharepoint

2. Download the current inventory list from the sharepoint

3. Open the anaconda prompt

4. run `conda activate SE_lottery`

3. specify the directries where you put those files at the top of the script (variable forms_dir)

4. the results will be written to a subdirectory of forms_dir. If you want to change that name, change the variable results_dir

5. change the dates and times for the last lottery and the deadline. Only applications in between those dates are used.

6. (only if the inventory numbers of ski changed) update the inventory numbers of skis

5. run `python PATH_TO_THE_LOTTERY_SCRIPT`


## What does it do?
It takes the applications and the inventory and lotteries items, that have more requests than there is in stock. 

## What's up with skis?
Skis are treated extra. If you get one kind of skis, you won't get a second pair. There is no lottery on ski boots. The reasoning behind this is, that there are more boots than skis. And people can fit multiple shoe sizes. So if they get a pair of skis, they can try to find a pair that fits them and they have a guaranteed pair of boots. If you would do lottery on the boots as well, you could end up in a situation, where you can get skis, but not boots