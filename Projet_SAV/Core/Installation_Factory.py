from Core.Installation_Definition import SolarInstallationFronius, SolarInstallationMC, SolarInstallationSMA, SolarInstallationVictron

class InstallationFactory:
    @staticmethod
    def create_installation(type_, name, id = None):
        if type_ == "victron energy":
            if id is None:
                raise ValueError("ID is required for Victron installations")
            return SolarInstallationVictron(name, id)
        elif type_ == "meteocontrol":
            return SolarInstallationMC(name)
        elif type_ == "SMA":
            return SolarInstallationSMA(name)
        elif type_ == "Fronius":
            return SolarInstallationFronius(name)
        else:
            raise ValueError(f"Unknown installation type: {type_}")
