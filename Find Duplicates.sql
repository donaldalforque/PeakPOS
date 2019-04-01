use peak
go

WITH Cte AS(
    SELECT *,
        Rn = ROW_NUMBER() OVER(PARTITION BY name ORDER BY ProductId DESC)
    FROM base_product
)
DELETE FROM Cte WHERE Rn > 1;

