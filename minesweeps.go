package main

import (
	"fmt"
	"os"
	"os/exec"
	"image"
	"image/color"
	"sort"
	"time"
)

var cnt int = 0
var failedToFindBombCell int = 0
var faildedToFindFreeCell int = 0
var a [][]int
var totolBombsFound = 0

func main() {
	exec.Command("gnome-mines", "--big").Start()
	fmt.Println("lets start")

	time.Sleep(3 * time.Second)

	// click on center cell
	clickLeftXY(24+38*15+19, 93+38*8+19)

	time.Sleep(1 * time.Second)

	a = make([][]int, 16)
	for i := range a {
		a[i] = make([]int, 30)
	}

	fillAllArr()

	iteration := 0
	for iteration < 200 {

		failedToFindBombCell++
		faildedToFindFreeCell++
        markBombCells()

		clickFreeCell()

		if totolBombsFound == 99 {
			fmt.Println("Solved it :D")
			os.Exit(0)
		}
		if failedToFindBombCell >= 4 && faildedToFindFreeCell >= 4 {

			unopenedCellArr := retUnpenedCellArr(a)

			allRegions := segregate(unopenedCellArr)

			allPoints := make([]struct {
				p   point
				cnt int
			}, 0)

			numOfRegionsChecked := 0
			for _, region := range allRegions {

				if len(region) > 26 {
					continue
				}
				numOfRegionsChecked++
				points := tankSolverAlgorithm(region, 0)

				for idx := range points {
					allPoints = append(allPoints, struct {
						p   point
						cnt int
					}{region[idx], points[idx]})
				}

			}
			if numOfRegionsChecked == 0 {
				fmt.Println("sorry i can't solve it because it will take long time :(")
				os.Exit(0)
			}
			sort.Slice(allPoints, func(i, j int) bool {
				return allPoints[i].cnt < allPoints[j].cnt
			})

			if len(allPoints) > 0 && allPoints[0].cnt > 0 {
				clickLeftXY(allPoints[0].p.Y*38+24+19, allPoints[0].p.X*38+19+93)
			} else {
				for _, point := range allPoints {
					if point.cnt == 0 {
						clickLeftXY(point.p.Y*38+24+19, point.p.X*38+19+93)
					} else {
						break
					}
				}
			}
			failedToFindBombCell = 0
			faildedToFindFreeCell = 0
		}
		fillAllArr()
		iteration++
	}

}

type cellColor struct {
	R int
	G int
	B int
}

func max(x, y int) int {
	if x < y {
		return y
	}
	return x
}

func getMax(a, b, c, d, e, f, g, h, i int) int {
	return max(a, max(b, max(c, max(d, max(e, max(f, max(g, max(h, i))))))))
}

func getCellNumber(img image.Image) int {
	unopenedCell := cellColor{186, 189, 182}
	zeroCell := cellColor{222, 222, 220}
	oneCell := cellColor{221, 250, 195}
	twoCell := cellColor{236, 237, 191}
	threeCell := cellColor{237, 218, 180}
	fourCell := cellColor{237, 195, 138}
	fiveCell := cellColor{247, 161, 162}
	sixCell := cellColor{254, 167, 133}
	sevenCell := cellColor{255, 125, 96}
	bombCell := cellColor{204, 0, 0}

	rect := img.Bounds()

	cellunopendCnt := 0
	cellZeroCnt := 0
	cellOneCnt := 0
	cellTwoCnt := 0
	cellThreeCnt := 0
	cellFourCnt := 0
	cellFiveCnt := 0
	cellSixCnt := 0
	cellSevenCnt := 0

	for i := 0; i < rect.Max.Y; i++ {
		for j := 0; j < rect.Max.X; j++ {
			c := color.RGBAModel.Convert(img.At(j, i))
			r := int(c.(color.RGBA).R)
			g := int(c.(color.RGBA).G)
			b := int(c.(color.RGBA).B)
			//fmt.Printf("%d %d %d \n", r, g, b)
			if bombCell.R == r && bombCell.B == b && bombCell.G == g {
				fmt.Println("Sorry Failed to solve it :(")
				os.Exit(0)
			}
			if unopenedCell.R == r && unopenedCell.B == b && unopenedCell.G == g {
				cellunopendCnt++
			}
			if zeroCell.R == r && zeroCell.B == b && zeroCell.G == g {
				cellZeroCnt++
			}
			if oneCell.R == r && oneCell.B == b && oneCell.G == g {
				cellOneCnt++
			}
			if twoCell.R == r && twoCell.B == b && twoCell.G == g {
				cellTwoCnt++
			}
			if threeCell.R == r && threeCell.B == b && threeCell.G == g {
				cellThreeCnt++
			}
			if fourCell.R == r && fourCell.B == b && fourCell.G == g {
				cellFourCnt++
			}
			if fiveCell.R == r && fiveCell.B == b && fiveCell.G == g {
				cellFiveCnt++
			}
			if sixCell.R == r && sixCell.B == b && sixCell.G == g {
				cellSixCnt++
			}
			if sevenCell.R == r && sevenCell.B == b && sevenCell.G == g {
				cellSevenCnt++
			}
		}
	}
	ret := -100
	if cellunopendCnt == getMax(cellunopendCnt, cellZeroCnt, cellOneCnt, cellTwoCnt, cellThreeCnt, cellFourCnt, cellFiveCnt, cellSixCnt, cellSevenCnt) {
		ret = 9
	}
	if cellZeroCnt == getMax(cellunopendCnt, cellZeroCnt, cellOneCnt, cellTwoCnt, cellThreeCnt, cellFourCnt, cellFiveCnt, cellSixCnt, cellSevenCnt) {
		ret = 0
	}
	if cellOneCnt == getMax(cellunopendCnt, cellZeroCnt, cellOneCnt, cellTwoCnt, cellThreeCnt, cellFourCnt, cellFiveCnt, cellSixCnt, cellSevenCnt) {
		ret = 1
	}
	if cellTwoCnt == getMax(cellunopendCnt, cellZeroCnt, cellOneCnt, cellTwoCnt, cellThreeCnt, cellFourCnt, cellFiveCnt, cellSixCnt, cellSevenCnt) {
		ret = 2
	}
	if cellThreeCnt == getMax(cellunopendCnt, cellZeroCnt, cellOneCnt, cellTwoCnt, cellThreeCnt, cellFourCnt, cellFiveCnt, cellSixCnt, cellSevenCnt) {
		ret = 3
	}
	if cellFourCnt == getMax(cellunopendCnt, cellZeroCnt, cellOneCnt, cellTwoCnt, cellThreeCnt, cellFourCnt, cellFiveCnt, cellSixCnt, cellSevenCnt) {
		ret = 4
	}
	if cellFiveCnt == getMax(cellunopendCnt, cellZeroCnt, cellOneCnt, cellTwoCnt, cellThreeCnt, cellFourCnt, cellFiveCnt, cellSixCnt, cellSevenCnt) {
		ret = 5
	}
	if cellSixCnt == getMax(cellunopendCnt, cellZeroCnt, cellOneCnt, cellTwoCnt, cellThreeCnt, cellFourCnt, cellFiveCnt, cellSixCnt, cellSevenCnt) {
		ret = 6
	}
	if cellSevenCnt == getMax(cellunopendCnt, cellZeroCnt, cellOneCnt, cellTwoCnt, cellThreeCnt, cellFourCnt, cellFiveCnt, cellSixCnt, cellSevenCnt) {
		ret = 7
	}
	if ret == -100 {
		fmt.Println("Failed to detect the screen")
		os.Exit(1)
	}

	return ret
}

func clickLeftXY(x, y int) {
	robotgo.MoveMouse(x, y)
	robotgo.MouseClick()
	robotgo.MoveMouse(-x, -y)
}

func retUnpenedCellArr(a [][]int) []point {
	var arr []point
	i := 0
	for i < 16 {
		j := 0
		for j < 30 {
			if a[i][j] == 9 {
				numberOfOpenedNeighbours := 0
				h := -1
				k := -1

				for h <= 1 {
				func fillAll	k = -1
					for k <= 1 {
						if k == 0 && h == 0 {
							k++
							continue
						}
						newX := i + h
						newY := j + k
						if newX < 16 && newX >= 0 && newY < 30 && newY >= 0 {
							if a[newX][newY] >= 0 && a[newX][newY] <= 7 {
								numberOfOpenedNeighbours++
							}
						}
						k++
					}
					h++
				}
				if numberOfOpenedNeighbours > 0 {
					arr = append(arr, point{i, j})
				}
			}
			j = j + 1
		}
		i = i + 1
	}
	return arr
}
