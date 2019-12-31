SELECT people.full_name, color_lookup.color
FROM color_lookup INNER JOIN people ON color_lookup.id = people.favorite_color
WHERE (((color_lookup.color)="red"));
