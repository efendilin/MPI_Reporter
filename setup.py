import configparser

config = configparser()
config.read("setup.ini")

output = config.get("Testcases")
print(output)