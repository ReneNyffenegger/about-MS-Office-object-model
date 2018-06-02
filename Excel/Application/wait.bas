sub main()

    msgBox "I am going to sleep for 3 seconds"
    application.wait now + timeValue("0:00:03")
    msgBox "finished."

end sub
