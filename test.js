let resizeAddress = resizeText(model.applicant_information.address, 180, BODY_SIZE);
        Shot('sf', 'tahoma', 'normal', resizeAddress);
        Shot("t", model.applicant_information.address, 20, 41.5, 'l')