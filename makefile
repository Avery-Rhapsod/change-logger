.PHONY: run

run:
	kitty --title "Server" sh -c "npm run server; exec sh" &
	kitty --title "Dev" sh -c "npm run dev; exec sh" &
