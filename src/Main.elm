module Main exposing (main)

import Browser
import Canvas exposing (..)
import Canvas.Settings exposing (..)
import Canvas.Settings.Line exposing (..)
import Color exposing (..)
import Html exposing (..)
import Http exposing (..)


type alias Model =
    {}


type Msg
    = NoOp


update : Msg -> Model -> ( Model, Cmd Msg )
update msg model =
    ( model, Cmd.none )


main : Program () Model Msg
main =
    Browser.element
        { init = \() -> ( {}, Cmd.none )
        , view = view
        , update = update
        , subscriptions = \model -> Sub.none
        }


view : Model -> Html Msg
view model =
    let
        width =
            1500

        height =
            900
    in
    
        Canvas.toHtml ( width, height )
            []
            [ shapes [ fill Color.red, stroke Color.yellow, lineWidth 15 ]
                [ circle ( 220, 220 ) 360
                ]
            , shapes [ fill Color.green ]
                [ rect ( 100, 100 ) 50 80
                , rect ( 0, 0 ) 40 30
                , circle ( 12, 12 ) 100
                , path ( 10, 300 )
                    [ lineTo ( 300, 200 )
                    , lineTo ( 30, 30 )
                    , lineTo ( 20, 10 )
                    ]
                ]
            , shapes [ stroke Color.blue, lineWidth 10 ]
                [ circle ( 300, 300 ) 50
                , circle ( 380, 300 ) 50
                , path ( 300, 300 )
                    [ lineTo ( 300, 500 )
                    , lineTo ( 320, 520 )
                    , lineTo ( 360, 520 )
                    , lineTo ( 380, 500 )
                    , lineTo ( 380, 300 )
                    , lineTo ( 300, 300 )
                    ]
                ]
            ]
    