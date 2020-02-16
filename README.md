# popo
Make Repeat Text From Exel

I made this project from an idea of "Automate the Boring Stuff with Python" of Al Sweigart. When coding a web using SpringBoot and JDBC, I often meet some kind of repeat text, which is thoundsands of line. Example:

    @Override
    @Transactional(rollbackFor = Exception.class)
    public Integer addNewUser(UserDto user) throws DataAccessException {
      String sql = "INSERT INTO database (" +
                ", user_id " +
                ", mail_address " +
                ", password " +
                ", param1 " +
                ", param2 " +
                ", param3 " +
                ", param4 " +
                .....
                ", param250) " +
                " VALUES ( " +
                ", :user_id" +
                ", :mail_address" +
                ", :password" +
                ", :param1" +
                ", :param2" +
                ", :param3" +
                ", :param4" +
                .....
                ", param250) "
                
                // About 500 lines
                Map<String, Object> params = new HashMap<String, Object>();

                params.put("user_id", user.getUser_id());
                params.put("mail_address", user.getMail_address());
                params.put("password", user.getPassword());
                params.put("param1", user.getParam1());
                .....
                // About 250 lines
                
                jdbc.update(sql, params);
                return 0;
    }

With this program, this crazy stuf can the sorten to:

        @Override
        @Transactional(rollbackFor = Exception.class)
        public Integer addNewUser(UserDto user) throws DataAccessException {
              String sql = "INSERT INTO database (" +
                    <loop $data = data.xlsx>
                    ", ${name} " +
                    </loop>
                    <loop $data = data.xlsx>
                    params.put("${name}", user.get${capitalize(name)}());
                    </loop>
                    jdbc.update(sql, params);
                    return 0;

                    //12 lines
 
 
This will make save hours of coding

