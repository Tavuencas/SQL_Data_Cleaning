Select * from dbo.Sheet1$

--Update the Sale Data format

ALTER TABLE Sheet1$
Add SaleDateConverted Date;
Update Sheet1$
SET SaleDateConverted = CONVERT(Date,SaleDate)

------------------------
--Populate Address data
Select * from dbo.Sheet1$
--Where PropertyAddress is null
order by ParcelID

Select a.ParcelID, a.PropertyAddress, b.ParcelID, b.PropertyAddress, ISNULL(a.PropertyAddress,b.PropertyAddress)
From dbo.Sheet1$ a
JOIN dbo.Sheet1$ b
	on a.ParcelID = b.ParcelID
	AND a.[UniqueID ] <> b.[UniqueID ]
Where a.PropertyAddress is null

Update a
SET PropertyAddress = ISNULL(a.PropertyAddress,b.PropertyAddress)
From dbo.Sheet1$ a
JOIN dbo.Sheet1$ b
	on a.ParcelID = b.ParcelID
	AND a.[UniqueID ] <> b.[UniqueID ]
Where a.PropertyAddress is null

------------------
-- Breaking out Address into Address, City, State

Select PropertyAddress
From dbo.Sheet1$

--Where PropertyAddress is null
--order by ParcelID
SELECT
SUBSTRING(PropertyAddress, 1, CHARINDEX(',', PropertyAddress) -1 ) as Address
, SUBSTRING(PropertyAddress, CHARINDEX(',', PropertyAddress) + 1 , LEN(PropertyAddress)) as Address
From dbo.Sheet1$


ALTER TABLE dbo.Sheet1$
Add PropertySplitAddress Nvarchar(255);
Update dbo.Sheet1$
SET PropertySplitAddress = SUBSTRING(PropertyAddress, 1, CHARINDEX(',', PropertyAddress) -1 )

ALTER TABLE dbo.Sheet1$
Add PropertySplitCity Nvarchar(255);
Update dbo.Sheet1$
SET PropertySplitCity = SUBSTRING(PropertyAddress, CHARINDEX(',', PropertyAddress) + 1 , LEN(PropertyAddress))

Select * from dbo.Sheet1$

Select OwnerAddress
From dbo.Sheet1$

Select
PARSENAME(REPLACE(OwnerAddress, ',', '.') , 3)
,PARSENAME(REPLACE(OwnerAddress, ',', '.') , 2)
,PARSENAME(REPLACE(OwnerAddress, ',', '.') , 1)
From dbo.Sheet1$

ALTER TABLE dbo.Sheet1$
Add OwnerSplitAddress Nvarchar(255);
Update dbo.Sheet1$
SET OwnerSplitAddress = PARSENAME(REPLACE(OwnerAddress, ',', '.') , 3)

ALTER TABLE dbo.Sheet1$
Add OwnerSplitCity Nvarchar(255);
Update dbo.Sheet1$
SET OwnerSplitCity = PARSENAME(REPLACE(OwnerAddress, ',', '.') , 2)

ALTER TABLE dbo.Sheet1$
Add OwnerSplitState Nvarchar(255);
Update dbo.Sheet1$
SET OwnerSplitState = PARSENAME(REPLACE(OwnerAddress, ',', '.') , 1)

Select * from dbo.Sheet1$

-----------------------------------
-- Remove duplicates
WITH RowNumCTE AS(
Select *,
	ROW_NUMBER() OVER (
	PARTITION BY ParcelID,
				 PropertyAddress,
				 SalePrice,
				 SaleDate,
				 LegalReference
				 ORDER BY
					UniqueID
					) row_num

From dbo.Sheet1$
--order by ParcelID
)
Select *
From RowNumCTE
Where row_num > 1
Order by PropertyAddress

Select * from dbo.Sheet1$

-----------------------------
-- Delete Unused Columns

Select * from dbo.Sheet1$

ALTER TABLE dbo.Sheet1$
DROP COLUMN OwnerAddress, TaxDistrict, PropertyAddress, SaleDate

Select * from dbo.Sheet1$